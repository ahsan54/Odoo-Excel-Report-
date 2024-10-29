# Odoo-Excel-Report-



## Model.Py

    def print_xlsx_salary_register(self):
        return self.env.ref("ModuleName.Report_Id").report_action(self)





## Report.Xml

<odoo>
    <report
            id="Report_Id"
            model="hr.payslip.run"
            report_type="xlsx"
            string="Salary Register"
            <!--            name="ModuleName.Report_Name"--> We Are not using this , delaing with direct file 
            file="ModuleName.Report_Name"
            menu="False"
    />
</odoo>





## Report.Py
rom odoo import fields, models, api, _

from odoo.exceptions import UserError

from odoo.tools import rgb_to_hex


class ClassName(models.AbstractModel):
    _name = "report.ModuleName.Report_Name"
    _inherit = 'report.report_xlsx.abstract'  # this is base module , keep it in directory


    def generate_xlsx_report(self, workbook, data, objs):
        payslip_ids = objs.slip_ids
        print(payslip_ids)
        sheet = workbook.add_worksheet(f"{objs.name}")

        bold_format = workbook.add_format({'bold': True})

        main_header_format = workbook.add_format({'bold': True, 'font_size': 55, 'align': 'center',
                                                  'valign': 'vcenter',
                                                  'border': 0
                                                  })

        header_format = workbook.add_format({
            'bold': True,
            # 'font_color': '#FFFFFF',  # White font
            # 'bg_color': '#0000FF',  # Blue background
            'font_size': 20,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,  # Adding border to cells
            'color': 'black',
            'text_wrap': True,
            'bg_color': "#D3D3D3",
        })

        data_format = workbook.add_format({
            'font_size': 20,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'color': 'purple',
            'text_wrap': True,
        })
        amount_format = workbook.add_format({             
            'font_size': 20,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'color': 'purple',
            'text_wrap': True,
            'num_format': '[$$]#,##0.00'
        })

        medium_format = workbook.add_format({
            'font_size' : 15 , 'align' : 'center' , 'valign': 'bottom'
        })# valign , yahi krta ha , ke text ko ider , uder uper nechay krta , center ma rakhte huwe.

        # Set the width of columns A to U to 30 (adjust the value as needed)
        sheet.set_column('A:U', 25)  # 30 is the width, adjust as per your need

        # Set the height of row 0 (first row) to 40 (adjust the value as needed)
        for rows in range(100):
            sheet.set_row(rows, 50)  # 50 is the height, adjust as per your need

        # we can merge cells by doing like this:-
        sheet.merge_range('A1:U1', 'Business Solutions & Services', main_header_format)
        sheet.merge_range('A3:U3', f'For The Month Of {str(objs.date_start.strftime("%B, %Y")).upper()}', header_format)

        sheet.write(3, 0, 'SR#', header_format)
        sheet.write(3, 1, 'Name', header_format)
        sheet.write(3, 2, 'Emp No.', header_format)
        sheet.write(3, 3, 'CNIC #', header_format)
        sheet.write(3, 4, 'Bank A/C #', header_format)
        sheet.write(3, 5, 'EOBI #', header_format)
        sheet.write(3, 6, 'Emp Joining Date #', header_format)
        sheet.write(3, 7, 'Designation', header_format)
        sheet.write(3, 8, 'Gross Salary', header_format)
        sheet.write(3, 9, 'Basic Salary', header_format)
        sheet.write(3, 10, 'Re-Imb', header_format)
        sheet.write(3, 11, 'Total', header_format)
        sheet.write(3, 12, 'Income Tax', header_format)
        sheet.write(3, 13, 'Health Ins.', header_format)
        sheet.write(3, 14, 'EOBI Ded', header_format)
        sheet.write(3, 15, 'Loan', header_format)
        sheet.write(3, 16, 'UnPaid Leaves', header_format)
        sheet.write(3, 17, 'Other Deduction', header_format)
        sheet.write(3, 18, 'Total Deduction', header_format)
        sheet.write(3, 19, f'{str(objs.date_start.strftime("%b-%y")).upper()} Salary MCB' , header_format)
        sheet.write(3, 20, 'Previous Salary MCB', header_format)

        row = 4
        sr_no = 1
        for payslip in payslip_ids:
            sheet.write(row, 0, sr_no, data_format)
            sheet.write(row, 1, payslip.employee_id.name or 'NONE',
                        data_format)
            sheet.write(row, 2, '', data_format)
            sheet.write(row, 3, payslip.employee_id.identification_id or '', data_format)
            sheet.write(row, 4, payslip.employee_id.bank_account_id.acc_number or '', data_format)
            sheet.write(row, 5, payslip.employee_id.eobi or '', data_format)
            sheet.write(row, 6, payslip.employee_id.first_contract_date.strftime('%d%m%y'), data_format)
            sheet.write(row, 7, payslip.employee_id.job_title or '', data_format)

            gross_salary = self.env['hr.payslip.line'].search([('slip_id','=',payslip.id),('code','=','GROSS')])
            gross_salary = sum(line.total for line in gross_salary)
            grand_gross = 0
            grand_gross += gross_salary
            sheet.write(row, 8, gross_salary or 0, amount_format)

            basic_salary = self.env['hr.payslip.line'].search([('slip_id','=',payslip.id),('code','=','BASIC')])
            basic_salary = sum(line.total for line in basic_salary)
            grand_basic = 0
            grand_basic +=basic_salary
            sheet.write(row, 9, basic_salary or 0, data_format)

            reimbursement = self.env['hr.payslip.line'].search([('slip_id','=',payslip.id),('code','=','REIMBURSEMENT')])
            reimbursement = sum(line.total for line in reimbursement)
            grand_reimbursement = 0
            grand_reimbursement += reimbursement
            sheet.write(row, 10, reimbursement or 0, amount_format)


            sheet.write(row, 11, 0, amount_format)


            income_tax_ded = self.env['hr.payslip.line'].search([('slip_id','=',payslip.id),('code','=','INTX')])
            income_tax_ded = sum(line.total for line in income_tax_ded)
            grand_income_tax_ded = 0
            grand_income_tax_ded += income_tax_ded
            sheet.write(row, 12,income_tax_ded or 0, amount_format)

            row += 1
            sr_no += 1

        workbook.close()




    
