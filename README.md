# Herbstfest order summary tool
We have an autumn festival (german Herbstfest) where we sell food for pickup and delivery.
The order tool is a form on a wordpress website which generates order emails.
This project provides two python scripts which create order summary information from these emails.

## config-template.json
Populate with data and save to config.json (todo make filename command line parameter)

## check_orders_herbstfest.py
It performs the following actions
`while true:
    get all emails
    process them into several pandas dataframes
    update html summary on Wordpress
    create Excel file with summary on filepath
`
The script runs until it is interrupted by CTRL+c.

## print_orders_herbstfest.py
This creates a single PDF document with individual pages per order. 
