"""
Build script v4: strip vbaProject.bin and convert the intermediate
Pump Hydraulics_WebDeploy_v1.xlsx (still technically xlsm-shaped inside,
just with a .xlsx extension) into a genuine macro-free xlsx package.

Steps:
  1. Delete xl/vbaProject.bin.
  2. Remove its Content_Types entries: the workbook Override content-type
     (macroEnabled.main+xml -> sheet.main+xml) and the vbaProject Default
     extension entry (if a bare 'bin' Default remains for printerSettings,
     leave that - only remove the vbaProject-specific one if it's an
     Override, not a Default; vbaProject.bin normally has no dedicated
     Override, it relies on the workbook-level macroEnabled type instead).
  3. Remove the vbaProject relationship from xl/_rels/workbook.xml.rels.
  4. Drop workbookPr's codeName attribute (VBA project identifier, meaningless
     without a VBA project and can trigger a repair prompt if left dangling).
  5. Re-save.
"""
import zipfile
import re

SRC = '/sessions/pensive-funny-keller/mnt/outputs/Pump Hydraulics_WebDeploy_v1.xlsx'
DST = '/sessions/pensive-funny-keller/mnt/outputs/Pump Hydraulics_WebDeploy_v2.xlsx'

MACRO_CT = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
PLAIN_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'


def main():
    with zipfile.ZipFile(SRC, 'r') as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}

    # ---- 1. delete vbaProject.bin ----
    removed = []
    for part in ['xl/vbaProject.bin', 'xl/vbaProjectSignature.bin']:
        if part in data:
            del data[part]
            removed.append(part)
    print('Deleted VBA parts:', removed)

    # ---- 2. Content_Types: flip workbook content-type, drop vbaProject override if present ----
    ct = data['[Content_Types].xml'].decode('utf-8')
    before = ct
    ct = ct.replace(
        f'<Override PartName="/xl/workbook.xml" ContentType="{MACRO_CT}"/>',
        f'<Override PartName="/xl/workbook.xml" ContentType="{PLAIN_CT}"/>'
    )
    print('workbook.xml content-type flipped:', ct != before)
    ct, n = re.subn(r'<Override PartName="/xl/vbaProject\.bin"[^>]*/>', '', ct)
    print('Removed vbaProject.bin override:', n)
    data['[Content_Types].xml'] = ct.encode('utf-8')

    # ---- 3. workbook.xml.rels: remove vbaProject relationship ----
    rels = data['xl/_rels/workbook.xml.rels'].decode('utf-8')
    rels, n = re.subn(r'<Relationship Id="rId\d+"[^>]*Target="vbaProject\.bin"[^>]*/>', '', rels)
    print('Removed vbaProject relationship from workbook.xml.rels:', n)
    data['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

    # ---- 4. workbook.xml: drop codeName, force full recalc on load ----
    wb = data['xl/workbook.xml'].decode('utf-8')
    wb, n = re.subn(r'\s*codeName="[^"]*"', '', wb)
    print('Removed codeName attr occurrences:', n)

    if '<calcPr' in wb:
        if 'fullCalcOnLoad' not in wb:
            wb = re.sub(r'<calcPr([^/]*)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
            print('Added fullCalcOnLoad="1" to calcPr')
    else:
        # insert a calcPr before </workbook> closing (rare fallback)
        wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
        print('Inserted new calcPr with fullCalcOnLoad="1"')

    data['xl/workbook.xml'] = wb.encode('utf-8')

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
