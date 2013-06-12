#! /usr/bin/python2.7

import argparse
from lxml import etree
import os
import xlwt


class AtiXML():

    def __init__(self, path):
        self.root = self.parse_xml(path)

    def parse_xml(self, path):
        tree = etree.parse(path)
        return tree.getroot()

    @property
    def datasources(self):
        return self.root.xpath("dataSources/dataSource")

    @property
    def primdocs(self):
        return self.root.xpath("primDocs/primDoc")

    @property
    def quotes(self):
        return self.root.xpath("primDocs/primDoc/quotations/q")

    def document_quotes(self, docid):
        return self.root.xpath("primDocs/primDoc[@id='%s']/quotations/q" % docid)

    @property
    def codes(self):
        return self.root.xpath("codes/code")

    @property
    def supercodes(self):
        return self.root.xpath("superCodes/superCode")

    @property
    def memos(self):
        return self.root.xpath("memos/memo")

    @property
    def codefams(self):
        return self.root.xpath("families/codeFamilies/codeFamily")

    @property
    def codelinks(self):
        return self.root.xpath("links/objectSegmentLinks/codings/iLink")

    @property
    def memolinks(self):
        return self.root.xpath("links/objectSegmentLinks/memoings/iLink")

    def writerow(self, sheet, rownum, values):
        for celnum, value in enumerate(values):
            sheet.write(rownum, celnum, value)

    def write_data_sources(self, sheet):
        headers = ['id', 'loc', 'mime', 'device', 'tf']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, ds in enumerate(self.datasources, start=1):
            values = [ds.get(att) for att in headers]
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def write_primary_documents(self, sheet):
        headers = ['name', 'id', 'loc', 'au', 'cDate', 'mDate', 'qIndex']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, pd in enumerate(self.primdocs, start=1):
            values = [pd.get(att) for att in headers]
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def write_quotes(self, sheet):
        headers = ['name', 'id', 'au', 'cDate', 'mDate', 'loc', 'doc_id', 'text']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        rownum = 0
        for doc in self.primdocs:
            docid = doc.get('id')
            for q in self.document_quotes(docid=docid):
                values = [q.get(att) for att in headers[:-2]]
                values.append(docid)
                values.append(self.quote_text(quote=q))
                rownum += 1
                self.writerow(sheet=sheet, rownum=rownum, values=values)

    def quote_text(self, quote):
        return ''.join([p.text for p in quote.findall('content/p') if p.text])

    def write_codes(self, sheet):
        headers = ['name', 'id', 'au', 'cDate', 'mDate', 'color', 'cCount', 'qCount']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, code in enumerate(self.codes, start=1):
            values = [code.get(att) for att in headers]
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def write_supercodes(self, sheet):
        pass

    def write_memos(self, sheet):
        headers = ['name', 'id', 'au', 'cDate', 'mDate', 'type', 'mime', 'fn', 'comments']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, memo in enumerate(self.memos, start=1):
            values = [memo.get(att) for att in headers[:-1]]
            values.append(self.memo_comments(memo=memo))
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def memo_comments(self, memo):
        return ' '.join([p.text for p in memo.findall('comment/p') if p.text])

    def write_code_families(self, sheet):
        headers = ['name', 'id', 'au', 'cDate', 'mDate']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, codefam in enumerate(self.codefams, start=1):
            values = [codefam.get(att) for att in headers]
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def write_code_family_members(self, sheet):
        headers = ['code_family_id', 'code_id']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        rownum = 0
        for codefam in self.codefams:
            famid = codefam.get('id')
            for item in codefam.findall("item"):
                rownum += 1
                values = [famid, item.get('id')]
                self.writerow(sheet=sheet, rownum=rownum, values=values)

    def write_code_links(self, sheet):
        headers = ['quote_id', 'code_id']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, link in enumerate(self.codelinks, start=1):
            values = [link.get('qRef'), link.get('obj')]
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def write_memo_links(self, sheet):
        headers = ['quote_id', 'memo_id']
        self.writerow(sheet=sheet, rownum=0, values=headers)
        for rownum, link in enumerate(self.memolinks, start=1):
            values = [link.get('qRef'), link.get('obj')]
            self.writerow(sheet=sheet, rownum=rownum, values=values)

    def export_to_excel(self, filename):
        wkbk = xlwt.Workbook(encoding='utf-8')

        datasources = wkbk.add_sheet('data sources')
        self.write_data_sources(sheet=datasources)

        primdocs = wkbk.add_sheet('primary documents')
        self.write_primary_documents(sheet=primdocs)

        quotes = wkbk.add_sheet('quotes')
        self.write_quotes(sheet=quotes)

        codes = wkbk.add_sheet('codes')
        self.write_codes(sheet=codes)

        supercodes = wkbk.add_sheet('supercodes')
        self.write_supercodes(sheet=supercodes)

        memos = wkbk.add_sheet('memos')
        self.write_memos(sheet=memos)

        codefams = wkbk.add_sheet('code families')
        self.write_code_families(sheet=codefams)

        members = wkbk.add_sheet('code family members')
        self.write_code_family_members(sheet=members)

        codelinks = wkbk.add_sheet('code links')
        self.write_code_links(sheet=codelinks)

        memolinks = wkbk.add_sheet('memo links')
        self.write_memo_links(sheet=memolinks)

        wkbk.save(filename)


def main():
    parser = argparse.ArgumentParser(description='Convert an Atlas.ti XML ' + \
        'dump to an Excel Spreadsheet')
    parser.add_argument('atlas', help='path to the Atlas.ti XML file')
    parser.add_argument('-e', '--excel', help='name of new Excel file.  ' + \
        'Defaults to same name as Atlas.ti file with xls extension.')
    args = parser.parse_args()
    if not getattr(args, 'excel', None):
        path, filename = os.path.split(args.atlas)
        args.excel = filename.replace('.xml', '.xls')
    ati = AtiXML(path=args.atlas)
    ati.export_to_excel(filename=args.excel)


if __name__ == '__main__':
    main()