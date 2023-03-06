# -*- coding: utf-8 -*-

from odoo import models
from datetime import datetime, timedelta, date

class ReportXlsx(models.AbstractModel):
    _name = "report.order_to_invoice.report_referral_order_xlsx"
    _inherit = "report.report_xlsx.abstract"

    def generate_xlsx_report(self, workbook, data, partners):
        format1 = workbook.add_format({'font_size': 13, 'align': 'vcenter', 'bold': True})
        format2 = workbook.add_format({'font_size': 10, 'align': 'center'})
        format3 = workbook.add_format({'font_size': 10, 'align': 'left'})
        format4 = workbook.add_format({'font_size': 10, 'align': 'right'})
        sheet = workbook.add_worksheet('Remisión #: ' + partners.folio)
        sheet.merge_range(0,0,0,12, '', format2)
        sheet.merge_range(1,0,1,12, '', format2)
        sheet.merge_range(2,2,2,6, 'INSTITUTO NACIONAL DE CARDIOLOGÍA "IGNACIO CHAVEZ"', format2)
        date = datetime.strftime(partners.date_order, '%d de %B de %Y')
        sheet.merge_range(2,8,2,10, date, format4)
        sheet.merge_range(3,2,3,6, 'JUAN BADIANO NO 1', format3)
        sheet.merge_range(3,8,3,9, 'SECCION XVI', format3)
        sheet.merge_range(4,2,4,6, 'TLALPAN, MEXICO D.F', format3)
        sheet.write(4,8, 14080, format3)
        sheet.merge_range(4,10,4,12, 'INC-430623-C16', format2)
        col_code = 1
        row = 5
        col = 3
        col_catalogue = 8
        col_quant = 10
        col_price = 11
        col_subtotal = 12
        subtotal_all = 0
        for line in partners.product_referral_ids:
            row += 1
            sheet.merge_range(row,col_code,row,2, line.code, format2)
            sheet.merge_range(row,col,row,7, line.product_id.name, format3)
            sheet.write(row,col_catalogue, line.catalog, format3)
            sheet.write(row,col_quant, line.quant, format2)
            sheet.write(row,col_price, line.product_id.list_price, format2)
            #Calculates subtotal in each line
            subtotal = line.quant * line.product_id.list_price
            sheet.write(row,col_subtotal, subtotal, format2)
            subtotal_all += subtotal
        sheet.write(16,12, subtotal_all, format3)
        #Calculates IVA
        subtotalIVA = subtotal_all * .16
        sheet.write(17,12, subtotalIVA, format3)
        #Calculates Total with IVA
        total = subtotal_all + subtotalIVA
        sheet.write(18,12, total, format3)
    