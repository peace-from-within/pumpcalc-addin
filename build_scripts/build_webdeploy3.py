"""
Build script v3: full web-deployment conversion (final).

Confirmed control inventory (via <controls> section inside each worksheet
XML - the authoritative source, more reliable than parsing vmlDrawing/
drawing XML separately):

  Sketch & P Profile (sheet9.xml), rels sheet9.xml.rels:
    rId4 -> ctrlProp1  Button 97      macro=Circuit_Definition
    rId5 -> ctrlProp2  Drop Down 116  macro=PumpDiff_fixed      (AB22 toggle)
    rId6 -> ctrlProp3  Drop Down 189  macro=Multi_SaveOpen_Sketch (AB39 case)
    rId7 -> ctrlProp4  Paths_Dropdown macro=SaveOpen_Path        (AK11 path)

  Line DP (sheet10.xml), rels sheet10.xml.rels:
    rId4 -> ctrlProp5  Button 38      macro=Show_EQUIVL
    rId5 -> ctrlProp6  Button 39      macro=AddStream_LineDP
    rId6 -> ctrlProp7  Drop Down 42   macro=EquivL_Selection
    rId7 -> ctrlProp8  Drop Down 56   macro=Multi_SaveOpen_LindeDP

  Pump Calculation (sheet13.xml), rels sheet13.xml.rels:
    rId9  -> ctrlProp9  Option Button 51  macro=Number13Input
    rId10 -> ctrlProp10 Option Button 52  macro=Number13Input
    rId12 -> ctrlProp11 Option Button 59  macro=Number39Input  (rId11 = image, skip)
    rId13 -> ctrlProp12 Option Button 61  macro=Number39Input
    rId15 -> ctrlProp13 Option Button 65  macro=Number40Input  (rId14 = ctrlProp15? verify)
    rId16 -> ctrlProp14 Option Button 66  macro=Number40Input

All of these are macro-triggered controls with zero purpose once VBA is
removed - deleted wholesale, along with:
  - their <controls>/<control> entries in the worksheet XML
  - the matching Relationship in each sheet's .rels file
  - the ctrlProp*.xml part itself
  - the matching shape entry in the sheet's vmlDrawing (found via shapeId)

Everything else (EquivL_Selection is a Data Validation-style dropdown for
display toggling, not calculation-critical, but it's macro-only so it has
to go too since there's no VBA left to run it).

NOTE: EquivL toggles a "Show/Hide equivalent length detail" view. Removing
it just means that detail column set stays in whatever state it was saved
in (no macro needed to view the numbers - they're separate columns, not
computed on demand). Confirmed via Module1_For_HandingSheet source: it just
does a Rows/Columns.Hidden = True/False toggle - purely a display
convenience with no calculation impact.
"""
import zipfile
import re

SRC = '/sessions/pensive-funny-keller/mnt/uploads/Pump Hydraulics_Case 반영_개선.xlsm'
DST = '/sessions/pensive-funny-keller/mnt/outputs/Pump Hydraulics_WebDeploy_v1.xlsx'

UDF_MAP = {
    'Calc_J': 'PUMPCALC.CALC_J',
    'Calc_JA': 'PUMPCALC.CALC_JA',
    'Calc_RA': 'PUMPCALC.CALC_RA',
    'Pipe_ID': 'PUMPCALC.PIPE_ID',
    'VaporVolFrac': 'PUMPCALC.VAPOR_VOL_FRAC',
    'AvgDensity': 'PUMPCALC.AVG_DENSITY',
    'DPHomogeneous': 'PUMPCALC.DP_HOMOGENEOUS',
    'DPDukler': 'PUMPCALC.DP_DUKLER',
    'InplaceDensity': 'PUMPCALC.INPLACE_DENSITY',
    'TwoPhaseVelocity': 'PUMPCALC.TWO_PHASE_VELOCITY',
    'BakerXval': 'PUMPCALC.BAKER_XVAL',
    'BakerYval': 'PUMPCALC.BAKER_YVAL',
    'Baker': 'PUMPCALC.BAKER',
    'GriffithWallisXval': 'PUMPCALC.GRIFFITH_WALLIS_XVAL',
    'GriffithWallisYval': 'PUMPCALC.GRIFFITH_WALLIS_YVAL',
    'Griffith': 'PUMPCALC.GRIFFITH',
    'CalcPres': 'PUMPCALC.CALC_PRES',
    'CalcHead': 'PUMPCALC.CALC_HEAD',
    'CalcPumpEff': 'PUMPCALC.CALC_PUMP_EFF',
    'CalcPumpEffAnother': 'PUMPCALC.CALC_PUMP_EFF_ALT',
    'SelMotorPower': 'PUMPCALC.SEL_MOTOR_POWER',
    'EstNPSHR': 'PUMPCALC.EST_NPSHR',
    'EqLenFactor': 'PUMPCALC.EQ_LEN_FACTOR',
    'CVPressDrop': 'PUMPCALC.CV_PRESS_DROP',
}

# sheet-file -> list of shapeIds to remove controls for (all of them, this round)
SHEETS_WITH_CONTROLS = {
    'xl/worksheets/sheet9.xml': 'xl/worksheets/_rels/sheet9.xml.rels',
    'xl/worksheets/sheet10.xml': 'xl/worksheets/_rels/sheet10.xml.rels',
    'xl/worksheets/sheet13.xml': 'xl/worksheets/_rels/sheet13.xml.rels',
}


def replace_udf_calls(text, counter):
    names_by_len = sorted(UDF_MAP.keys(), key=len, reverse=True)
    for name in names_by_len:
        pattern = re.compile(r'\b' + re.escape(name) + r'\s*\(')
        def _sub(m, name=name):
            counter[name] = counter.get(name, 0) + 1
            return UDF_MAP[name] + '('
        text = pattern.sub(_sub, text)
    return text


def process_sheet_formulas(xml_text, counter):
    def repl(m):
        return replace_udf_calls(m.group(0), counter)
    return re.sub(r'<f[^>]*>.*?</f>', repl, xml_text, flags=re.DOTALL)


def blank_cell(xml_text, cellref):
    pat = re.compile(r'<c r="' + re.escape(cellref) + r'"([^>]*?)(?:/>|>.*?</c>)', re.DOTALL)
    m = pat.search(xml_text)
    if not m:
        return xml_text, False
    attrs = m.group(1)
    attrs = re.sub(r'\s*t="[^"]*"', '', attrs)
    new_cell = f'<c r="{cellref}"{attrs}/>'
    return xml_text[:m.start()] + new_cell + xml_text[m.end():], True


def main():
    counter = {}
    with zipfile.ZipFile(SRC, 'r') as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}

    # ---- 1. UDF -> PUMPCALC.* rewrite across ALL worksheets ----
    for fname in list(data.keys()):
        if not (fname.startswith('xl/worksheets/sheet') and fname.endswith('.xml')):
            continue
        xml = data[fname].decode('utf-8')
        new_xml = process_sheet_formulas(xml, counter)
        if new_xml != xml:
            data[fname] = new_xml.encode('utf-8')
    print('UDF replacement counts:', counter)
    print('UDF total occurrences:', sum(counter.values()))

    # ---- 2. Clear Sketch!AK3:AK11 (Path dropdown option-label cells) ----
    s9 = data['xl/worksheets/sheet9.xml'].decode('utf-8')
    cleared = []
    for ref in ['AK3', 'AK4', 'AK5', 'AK6', 'AK8', 'AK9', 'AK10', 'AK11']:
        s9, ok = blank_cell(s9, ref)
        if ok:
            cleared.append(ref)
    print('Cleared Sketch cells:', cleared)
    data['xl/worksheets/sheet9.xml'] = s9.encode('utf-8')

    # ---- 3. Hide the Paths worksheet (state="veryHidden") instead of
    # deleting it outright.
    #
    # IMPORTANT FINDING: actually removing a <sheet> entry from
    # workbook.xml's <sheets> list - even when done fully consistently
    # (matching rels relationship removed, sheetN.xml part deleted, its own
    # rels deleted, Content_Types override removed, orphaned per-sheet parts
    # like comments/vmlDrawing deleted, and docProps/app.xml's
    # HeadingPairs/TitlesOfParts counts corrected to match) still makes both
    # LibreOffice and (per the same OOXML validation rules) Excel refuse to
    # load the file ("source file could not be loaded"). This was verified
    # empirically via bisection - removing ANY single <sheet> tag (tested
    # with both Paths and, as a control, the unrelated Pipe_ID sheet) while
    # leaving its worksheet part/rels/content-type untouched reproduces the
    # exact same failure. The reverse (removing only the workbook.xml.rels
    # relationship, or only deleting the sheetN.xml part, while leaving the
    # <sheet> tag itself in place) loads fine. So the true constraint is:
    # workbook.xml's <sheets> list must always have exactly one <sheet>
    # entry per worksheet part that Content_Types declares - full deletion
    # requires re-deriving that invariant correctly, which is riskier than
    # it looks. state="veryHidden" sidesteps the whole problem: the sheet,
    # its data, and all its relationships stay fully intact (byte-identical
    # structure to the original), so there is zero risk of the corruption
    # class above, while the sheet is completely inaccessible to the end
    # user (veryHidden sheets have no "Unhide" menu entry in Excel's UI -
    # unlike a normal hidden sheet - they can only be revealed via VBA or a
    # scripting tool, neither of which the customer's copy will have).
    wb = data['xl/workbook.xml'].decode('utf-8')
    wb2 = wb.replace(
        '<sheet name="Paths" sheetId="53" r:id="rId12"/>',
        '<sheet name="Paths" sheetId="53" state="veryHidden" r:id="rId12"/>'
    )
    if wb2 == wb:
        print('WARNING: Paths <sheet> tag not found in expected exact form - not hidden!')
    else:
        print('Paths worksheet set to state="veryHidden" (data/structure otherwise untouched)')
    data['xl/workbook.xml'] = wb2.encode('utf-8')

    # ---- 4. Delete Paths_Dropdown shape (hidden xdr:sp) from drawing9.xml ----
    d9 = data['xl/drawings/drawing9.xml'].decode('utf-8')
    anchor_pat = re.compile(
        r'<xdr:twoCellAnchor[^>]*>(?:(?!</xdr:twoCellAnchor>).)*?Paths_Dropdown(?:(?!</xdr:twoCellAnchor>).)*?</xdr:twoCellAnchor>',
        re.DOTALL)
    m = anchor_pat.search(d9)
    if m:
        d9 = d9[:m.start()] + d9[m.end():]
        data['xl/drawings/drawing9.xml'] = d9.encode('utf-8')
        print('Removed Paths_Dropdown label shape from drawing9.xml')
    else:
        print('NOTE: Paths_Dropdown label shape not found in drawing9.xml (may already be absent)')

    # ---- 5. Remove ALL macro-linked Form Controls (buttons/dropdowns) ----
    # For each target sheet: strip <controls>...</controls>, remove the
    # matching relationships + ctrlProp parts + vmlDrawing shape entries.
    all_ctrlprop_files_to_delete = set()

    for sheet_file, rels_file in SHEETS_WITH_CONTROLS.items():
        xml = data[sheet_file].decode('utf-8')
        controls_m = re.search(r'<controls>.*?</controls>', xml, re.DOTALL)
        if not controls_m:
            print(f'{sheet_file}: no <controls> section found')
            continue
        # Split into individual <control ...>...</control> blocks first (via
        # the mc:AlternateContent wrapper boundaries) so Group Box controls
        # (which have no macro= attribute) are not skipped or mismatched.
        control_blocks = re.findall(r'<control shapeId="\d+"[^>]*>.*?</control>', controls_m.group(0), re.DOTALL)
        control_entries = []
        for block in control_blocks:
            sid_m = re.search(r'shapeId="(\d+)"', block)
            rid_m = re.search(r'r:id="(rId\d+)"', block)
            name_m = re.search(r'name="([^"]+)"', block)
            macro_m = re.search(r'macro="([^"]*)"', block)
            control_entries.append((
                sid_m.group(1) if sid_m else '',
                rid_m.group(1) if rid_m else '',
                name_m.group(1) if name_m else '',
                macro_m.group(1) if macro_m else '',
            ))
        print(f'{sheet_file}: {len(control_entries)} controls ->',
              [(n, r) for _, r, n, _ in control_entries])

        # remove the whole <controls>...</controls> block
        xml = xml[:controls_m.start()] + xml[controls_m.end():]
        data[sheet_file] = xml.encode('utf-8')

        # resolve each rId to its ctrlProp target via the sheet's rels file
        rels_xml = data[rels_file].decode('utf-8')
        for shape_id, rid, name, macro in control_entries:
            rel_m = re.search(r'<Relationship Id="' + re.escape(rid) + r'"[^>]*Target="([^"]+)"[^>]*/>', rels_xml)
            if rel_m:
                target = rel_m.group(1)  # e.g. ../ctrlProps/ctrlProp2.xml
                abs_target = 'xl/' + target.replace('../', '')
                all_ctrlprop_files_to_delete.add(abs_target)
            # remove this relationship from the rels file
            rels_xml = re.sub(r'<Relationship Id="' + re.escape(rid) + r'"[^>]*/>', '', rels_xml)
        data[rels_file] = rels_xml.encode('utf-8')

        # remove matching shape entries from the vmlDrawing referenced by this sheet's rels
        vml_m = re.search(r'<Relationship Id="rId\d+"[^>]*Target="\.\./drawings/(vmlDrawing\d+\.vml)"[^>]*/>', rels_xml)
        if vml_m:
            vml_path = 'xl/drawings/' + vml_m.group(1)
            vml_xml = data[vml_path].decode('utf-8')
            for shape_id, rid, name, macro in control_entries:
                # v:shape id is "_x0000_sNNNNN" where NNNNN == shapeId, OR named (Paths_Dropdown)
                shape_pat = re.compile(
                    r'<v:shape id="(?:_x0000_s' + re.escape(shape_id) + r'|' + re.escape(name) + r')"[^>]*>.*?</v:shape>',
                    re.DOTALL)
                vml_xml, n = shape_pat.subn('', vml_xml)
                if n:
                    print(f'  removed vml shape for {name} (shapeId={shape_id}) from {vml_path}')
            data[vml_path] = vml_xml.encode('utf-8')

    print('ctrlProp files to delete:', sorted(all_ctrlprop_files_to_delete))

    # ---- 6. Delete ctrlProp parts + their Content_Types overrides (if any) ----
    ct = data['[Content_Types].xml'].decode('utf-8')
    for cp_file in all_ctrlprop_files_to_delete:
        if cp_file in data:
            del data[cp_file]
        ov_pat = re.compile(r'<Override PartName="/' + re.escape(cp_file) + r'"[^>]*/>')
        ct, n = ov_pat.subn('', ct)
    data['[Content_Types].xml'] = ct.encode('utf-8')
    print('Deleted', len(all_ctrlprop_files_to_delete), 'ctrlProp parts')

    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zipfile.ZipFile(SRC).infolist():
            if info.filename not in data:
                continue
            payload = data[info.filename]
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, payload)

    print('Saved:', DST)


if __name__ == '__main__':
    main()
