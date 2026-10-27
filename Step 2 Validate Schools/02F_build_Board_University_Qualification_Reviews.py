"""Copy the current evaluation checklist to the latest cycle Review tab, preserving other tabs."""
from copy import deepcopy
from datetime import datetime
import argparse
import os
import re
import shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--folder', type=Path, default=Path(__file__).resolve().parent)
args = parser.parse_args()
folder = args.folder.resolve()
original = folder / 'University Qualification Board Reviews.xlsx'
authored = folder / 'University Qualification Tracking requirements.xlsx'
schedule_folder = folder.parent / 'Step 1  Generate Scholarship Dates and Request for Nominations'
pattern = re.compile(r'^(\d{4}) PHISIM Scholarship Process(?:ing)? Dates\.(?:docx|xlsx)$', re.I)
years = [int(match[1]) for path in schedule_folder.iterdir() if (match := pattern.fullmatch(path.name))]
if not years:
    raise ValueError('No annual scholarship schedule found; no workbook changed.')
year = max(years)
review_name = f'{year} Review'
result = folder / '.board-reviews-updating.xlsx'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
tag = lambda name: '{' + MAIN + '}' + name
dump = lambda root: E.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

with ZipFile(original) as old, ZipFile(authored) as new:
    workbook = E.fromstring(old.read('xl/workbook.xml'))
    relationships = E.fromstring(old.read('xl/_rels/workbook.xml.rels'))
    types = E.fromstring(old.read('[Content_Types].xml'))
    styles = E.fromstring(old.read('xl/styles.xml'))
    added_styles = E.fromstring(new.read('xl/styles.xml'))
    new_workbook = E.fromstring(new.read('xl/workbook.xml'))
    new_rels = E.fromstring(new.read('xl/_rels/workbook.xml.rels'))
    new_sheet = next(x for x in new_workbook.find(tag('sheets')) if x.get('name') == 'Current evaluation checklist')
    relation = next(x for x in new_rels if x.get('Id') == new_sheet.get('{' + REL + '}id'))
    sheet_path = relation.get('Target')
    sheet_path = sheet_path.lstrip('/') if sheet_path.startswith('/') else 'xl/' + sheet_path
    sheet = E.fromstring(new.read(sheet_path))
    sheet_rel = str(Path(sheet_path).parent).replace('\\', '/') + '/_rels/' + Path(sheet_path).name + '.rels'
    if sheet_rel in new.namelist():
        raise ValueError('Checklist has linked worksheet objects; no workbook changed.')
    if sheet.find(tag('conditionalFormatting')) is not None:
        raise ValueError('Checklist has conditional formatting requiring style mapping; no workbook changed.')
    strings = E.fromstring(new.read('xl/sharedStrings.xml')) if 'xl/sharedStrings.xml' in new.namelist() else []
    # Use inline strings to avoid changing the original shared-string table.
    for cell in sheet.iter(tag('c')):
        if cell.get('t') == 's':
            value = cell.find(tag('v'))
            item = strings[int(value.text)]
            cell.remove(value)
            inline = E.SubElement(cell, tag('is'))
            for child in item:
                inline.append(deepcopy(child))
            cell.set('t', 'inlineStr')
    maps = {}
    # Append styles used by the new tab, retaining every original style index.
    original_nums = styles.find(tag('numFmts'))
    added_nums = added_styles.find(tag('numFmts'))
    number_map = {}
    if added_nums is not None:
        if original_nums is None:
            original_nums = E.Element(tag('numFmts'))
            styles.insert(0, original_nums)
        next_id = max([163] + [int(x.get('numFmtId')) for x in original_nums]) + 1
        for entry in added_nums:
            number_map[int(entry.get('numFmtId'))] = next_id
            item = deepcopy(entry); item.set('numFmtId', str(next_id))
            original_nums.append(item); next_id += 1
        original_nums.set('count', str(len(original_nums)))
    for collection in ('fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs'):
        destination = styles.find(tag(collection))
        source = added_styles.find(tag(collection))
        if source is None:
            continue
        offset = len(destination)
        maps[collection] = {i: offset + i for i in range(len(source))}
        for entry in source:
            item = deepcopy(entry)
            if collection in ('cellStyleXfs', 'cellXfs'):
                for attribute, lookup in (('fontId','fonts'),('fillId','fills'),('borderId','borders'),('xfId','cellStyleXfs')):
                    if item.get(attribute) is not None and lookup in maps:
                        item.set(attribute, str(maps[lookup][int(item.get(attribute))]))
                num = int(item.get('numFmtId', '0'))
                if num in number_map:
                    item.set('numFmtId', str(number_map[num]))
            destination.append(item)
        destination.set('count', str(len(destination)))
    for cell in sheet.iter(tag('c')):
        cell.set('s', str(maps['cellXfs'][int(cell.get('s', '0'))]))
    sheets = workbook.find(tag('sheets'))
    next_sheet = max(int(x.get('sheetId')) for x in sheets) + 1
    next_rel = 'rId' + str(max(int(x.get('Id')[3:]) for x in relationships if x.get('Id','').startswith('rId')) + 1)
    path = f'xl/worksheets/qualification_review_{year}.xml'
    existing = next((x for x in sheets if x.get('name') == review_name), None)
    if existing is None:
        existing = next((x for x in sheets if x.get('name') == 'Board Members'), None)
    removed_path = None
    if existing is not None:
        existing_rid = existing.get('{' + REL + '}id')
        existing_rel = next(x for x in relationships if x.get('Id') == existing_rid)
        removed_path = existing_rel.get('Target')
        removed_path = removed_path.lstrip('/') if removed_path.startswith('/') else 'xl/' + removed_path
        sheets.remove(existing)
        relationships.remove(existing_rel)
        for item in list(types):
            if item.get('PartName') == '/' + removed_path:
                types.remove(item)
    E.SubElement(sheets, tag('sheet'), name=review_name, sheetId=str(next_sheet), attrib={'{' + REL + '}id': next_rel})
    E.SubElement(relationships, '{' + PKG + '}Relationship', Id=next_rel, Type=REL + '/worksheet', Target=f'worksheets/qualification_review_{year}.xml')
    E.SubElement(types, '{' + CT + '}Override', PartName='/' + path, ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    replaced = {'xl/workbook.xml':dump(workbook), 'xl/_rels/workbook.xml.rels':dump(relationships),
                '[Content_Types].xml':dump(types), 'xl/styles.xml':dump(styles)}
    with ZipFile(result, 'w', ZIP_DEFLATED) as output:
        for entry in old.infolist():
            if entry.filename == removed_path:
                continue
            output.writestr(entry, replaced.get(entry.filename, old.read(entry.filename)))
        output.writestr(path, dump(sheet))
# Validate the result before replacing the workbook. Preserve a dated backup.
from openpyxl import load_workbook
source_book = load_workbook(authored, data_only=False)
check_book = load_workbook(result, data_only=False)
try:
    source_sheet = source_book['Current evaluation checklist']
    copied_sheet = check_book[review_name]
    if str(year) not in str(source_sheet['A1'].value):
        raise ValueError('Checklist year does not match the latest schedule; no workbook changed.')
    for row in source_sheet:
        for cell in row:
            actual = copied_sheet[cell.coordinate]
            if cell.value != actual.value or cell.number_format != actual.number_format:
                raise ValueError(f'Copy verification failed at {cell.coordinate}; no workbook changed.')
finally:
    source_book.close()
    check_book.close()
backup = original.with_name(original.stem + '.backup-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f') + '.xlsx')
shutil.copy2(original, backup)
os.replace(result, original)
print(f'Saved {original} — {review_name}')
print(f'Backup: {backup}')
