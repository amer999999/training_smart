from datetime import date 
leases = env['real_estate.lease'].search([('state', '=', 'active'), ('end_date', '<', date.today()), ])
for lease in leases:
        print(f"lease name: {lease.name} will expire on: {lease.end_date}")
