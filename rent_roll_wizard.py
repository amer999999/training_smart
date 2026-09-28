# -*- coding: utf-8 -*-
import io
import base64
import xlsxwriter
from odoo import models, fields, api
from datetime import datetime

class RentRollWizard(models.TransientModel):
    _name = 'rent.roll.wizard'
    _description = 'Rent Roll Report Wizard'
    
    date_from = fields.Date(
        string='Date From',
        required=True,
        default=fields.Date.today
    )
    date_to = fields.Date(
        string='Date To',
        required=True,
        default=fields.Date.today
    )
    property_type = fields.Selection([
        ('apartment', 'Apartment'),
        ('house', 'House'),
        ('villa', 'Villa'),
        ('commercial', 'Commercial'),
    ], string='Property Type')
    
    state = fields.Selection([
        ('available', 'Available'),
        ('rented', 'Rented'),
    ], string='Status')
    
    include_vacant = fields.Boolean(
        string='Include Vacant Properties',
        default=True
    )
    
    excel_file = fields.Binary(string='Excel File', readonly=True)
    excel_filename = fields.Char(string='Filename', readonly=True)
    
    def action_generate_report(self):
        """Generate Excel rent roll report"""
        self.ensure_one()
        
        # Build domain for filtering
        domain = []
        
        if self.property_type:
            domain.append(('property_type', '=', self.property_type))
        
        if self.state:
            if self.state == 'available':
                domain.append(('available', '=', True))
            elif self.state == 'rented':
                domain.append(('available', '=', False))
        
        # Get properties
        properties = self.env['real_estate.property'].search(domain)
        
        # Get active leases within date range
        lease_domain = [
            ('start_date', '<=', self.date_to),
            '|',
            ('end_date', '>=', self.date_from),
            ('end_date', '=', False),
        ]
        
        if not self.include_vacant:
            lease_domain.append(('state', '=', 'active'))
        
        leases = self.env['real_estate.lease'].search(lease_domain)
        
        # Generate Excel
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        # Formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#4472C4',
            'font_color': 'white'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        category_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFC000',
            'border': 1,
            'font_size': 11
        })
        
        data_format = workbook.add_format({'border': 1})
        
        currency_format = workbook.add_format({
            'num_format': '$#,##0.00',
            'border': 1
        })
        
        date_format = workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'border': 1
        })
        
        total_format = workbook.add_format({
            'bold': True,
            'bg_color': '#E2EFDA',
            'num_format': '$#,##0.00',
            'border': 1
        })
        
        label_format = workbook.add_format({
            'bold': True,
            'border': 1
        })
        
        # Create worksheet
        worksheet = workbook.add_worksheet('Rent Roll')
        
        # Set column widths
        worksheet.set_column('A:A', 25)  # Property Name
        worksheet.set_column('B:B', 15)  # Type
        worksheet.set_column('C:C', 20)  # Tenant
        worksheet.set_column('D:D', 15)  # Monthly Rent
        worksheet.set_column('E:E', 12)  # Start Date
        worksheet.set_column('F:F', 12)  # End Date
        worksheet.set_column('G:G', 10)  # Status
        
        # Title row
        row = 0
        worksheet.merge_range(row, 0, row, 6, 
                            f'Rent Roll Report ({self.date_from} to {self.date_to})', 
                            title_format)
        worksheet.set_row(row, 25)
        row += 2
        
        # Headers
        headers = ['Property Name', 'Type', 'Tenant', 'Monthly Rent', 
                  'Lease Start', 'Lease End', 'Status']
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, header_format)
        
        worksheet.set_row(row, 30)
        row += 1
        
        # Group properties by type
        property_types = properties.mapped('property_type')
        property_groups = {pt: properties.filtered(lambda p: p.property_type == pt) 
                          for pt in set(property_types)}
        
        grand_total = 0
        
        for prop_type, props in sorted(property_groups.items()):
            # Category header
            type_label = dict(self.env['real_estate.property']._fields['property_type'].selection).get(prop_type, prop_type)
            worksheet.write(row, 0, f'{type_label.upper()}', category_format)
            worksheet.merge_range(row, 1, row, 6, '', category_format)
            row += 1
            
            subtotal = 0
            
            for prop in props.sorted(key=lambda p: p.name):
                # Find active lease for this property
                prop_lease = leases.filtered(lambda l: l.property_id == prop and l.state == 'active')[:1]
                
                if prop_lease:
                    worksheet.write(row, 0, prop.name, data_format)
                    worksheet.write(row, 1, type_label, data_format)
                    worksheet.write(row, 2, prop_lease.tenant_id.name, data_format)
                    worksheet.write(row, 3, prop_lease.monthly_rent, currency_format)
                    worksheet.write(row, 4, prop_lease.start_date, date_format)
                    worksheet.write(row, 5, prop_lease.end_date or '', date_format)
                    worksheet.write(row, 6, 'Active', data_format)
                    
                    subtotal += prop_lease.monthly_rent
                    row += 1
                elif self.include_vacant:
                    worksheet.write(row, 0, prop.name, data_format)
                    worksheet.write(row, 1, type_label, data_format)
                    worksheet.write(row, 2, 'VACANT', data_format)
                    worksheet.write(row, 3, 0, currency_format)
                    worksheet.write(row, 4, '', data_format)
                    worksheet.write(row, 5, '', data_format)
                    worksheet.write(row, 6, 'Available', data_format)
                    row += 1
            
            # Subtotal row
            worksheet.write(row, 2, f'{type_label} Subtotal:', total_format)
            worksheet.write(row, 3, subtotal, total_format)
            worksheet.merge_range(row, 4, row, 6, '', total_format)
            row += 1
            
            grand_total += subtotal
        
        # Grand total row
        row += 1
        grand_total_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#70AD47',
            'font_color': 'white',
            'num_format': '$#,##0.00',
            'border': 2
        })
        
        worksheet.write(row, 2, 'GRAND TOTAL:', grand_total_format)
        worksheet.write(row, 3, grand_total, grand_total_format)
        worksheet.merge_range(row, 4, row, 6, '', grand_total_format)
        
        # Add summary statistics
        row += 3
        stats_header = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'bg_color': '#D9E1F2'
        })
        
        worksheet.write(row, 0, 'SUMMARY STATISTICS', stats_header)
        row += 1
        
        total_properties = len(properties)
        occupied = len(leases.filtered(lambda l: l.property_id in properties and l.state == 'active'))
        occupancy_rate = (occupied / total_properties * 100) if total_properties > 0 else 0
        
        worksheet.write(row, 0, 'Total Properties:', label_format)
        worksheet.write(row, 1, total_properties, data_format)
        row += 1
        
        worksheet.write(row, 0, 'Occupied:', label_format)
        worksheet.write(row, 1, occupied, data_format)
        row += 1
        
        worksheet.write(row, 0, 'Vacant:', label_format)
        worksheet.write(row, 1, total_properties - occupied, data_format)
        row += 1
        
        worksheet.write(row, 0, 'Occupancy Rate:', label_format)
        percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
        worksheet.write(row, 1, occupancy_rate / 100, percent_format)
        
        workbook.close()
        
        # Save to wizard
        output.seek(0)
        self.excel_file = base64.b64encode(output.read())
        self.excel_filename = f'Rent_Roll_{self.date_from}_{self.date_to}.xlsx'
        
        # Return action to download
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'rent.roll.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'views': [(False, 'form')],
            'target': 'new',
        }