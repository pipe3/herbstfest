#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# PROGRAMMER: oliver@kuhles.net
# DATE CREATED: 08/Oct/2021
# REVISED DATE:             <=(Date Revised - if any)
# PURPOSE: Herbstfest print order tool
#          To read formatted order emails from IMAP and summarize
#          them. Create a PDF with individual pages for each order
#          to be printed out and used during delivery
#          This tool is highly customized to our Herbstfest :-)

# Import stuff
from imbox import Imbox
from html_table_parser.parser import HTMLTableParser
import pandas as pd
from pandas import ExcelWriter
import numpy as np
from datetime import datetime
from time import time, sleep
import requests
import json
import base64
import sys
# import in SimpleDocTemplate
from reportlab.platypus import SimpleDocTemplate, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from tabulate import tabulate


# Retrieve orders from IMAP and process them
def get_bestellungen(server, user, psw, subject, folder, pdffile):

    with Imbox(server, user, psw,
        ssl=True,
        ssl_context=None,
        starttls=False) as imbox:

        # initialize vars
        bestellungen = []

        # Get all Herbstfest messages
        messages = imbox.messages(folder=folder, subject=subject)

        content = []

        styles = getSampleStyleSheet()
        styleT = styles['Title']
        styleN = styles['Normal']
        styleH1 = styles['Heading1']
        styleH1 = styles['Heading2']

        umsatz = 0

        for uid, message in messages:

            # instantiate the parser and feed it
            p = HTMLTableParser()
            p.feed(message.body['plain'][0])

            # our data is sub-table 0
            data = p.tables[0]

            page = []

            # Abholung oder Lieferung
            page.append("")
            if(data[9][1]).startswith('Abholung'):
                page.append('Abholung')
            elif(data[9][1]).startswith('Lieferung'):
                page.append('Lieferung')

            # Tag und Zeit
            if((data[10][1]).startswith('Samstag')):
                page.append('Samstag, '+data[11][1])
            elif((data[10][1]).startswith('Sonntag')):
                page.append('Sonntag, '+data[11][1])
            page.append("")

            # Name und Adresse
            page.append(data[13][1]+" "+data[14][1])
            if(data[9][1] == 'Lieferung'):
                page.append(data[17][1])
            page.append("Tel: "+data[16][1])
            page.append("")
            page.append("")
            page.append("")

            # Art
            page.append(data[2][1])
            page.append("")
            # Produkte
            page.append("Stück Zwiebel: "+data[3][1])
            page.append("Blech Zwiebel: "+data[4][1])
            page.append("Stück Flamm: "+data[5][1])
            page.append("Blech Flamm: "+data[6][1])
            # Neuer Suesser
            if(data[7][1].startswith('keiner')):
                page.append("Neuer Süßer: 0")
            elif (data[7][1].startswith('1')):
                page.append("Neuer Süßer: 1L")
            elif(data[7][1].startswith('2')):
                page.append("Neuer Süßer: 2L")
            elif(data[7][1].startswith('3')):
                page.append("Neuer Süßer: 3L")
            page.append("")

            # Quanta costa
            gesamtpreis = 0
            gesamtpreis += int(data[3][1])*2.5 # Stück Zwiebel
            gesamtpreis += int(data[4][1])*20 # Blech Zwiebel
            gesamtpreis += int(data[5][1])*2.5 # Stück Flamm
            gesamtpreis += int(data[6][1])*20 # Blech Flamm
            # Neuer Suesser
            if(data[7][1].startswith('keiner')):
                pass
            elif (data[7][1].startswith('1')):
                gesamtpreis += 7 # 1L
            elif(data[7][1].startswith('2')):
                gesamtpreis += 12 # 2L
            elif(data[7][1].startswith('3')):
                gesamtpreis += 18 # 3L

            umsatz += gesamtpreis

            gesamtpreis = "{:,.2f} EUR".format(gesamtpreis)

            page.append("Gesamt: "+gesamtpreis)


            content.append(Paragraph("Bestellung Herbstfest 2021", styleT))
            for p in page:
                content.append(Paragraph(p, styleH1))
            content.append(PageBreak())

        content.append(Paragraph("Umsatz: "+str(umsatz), styleH1))
        create_pdf(pdffile, content)


def create_pdf(filename, data):

    doc = SimpleDocTemplate(filename)
    doc.build(data)


def main():
    # Load the config file
    with open('config.json', 'r') as f:
        config = json.load(f)

    # Get vars from config file
    pdffile = config.get('pdffile')

    # IMAP
    imap_server = config.get('imap_server')
    imap_user = config.get('imap_user')
    imap_psw = config.get('imap_psw')
    imap_subject = config.get('imap_subject')
    imap_folder = config.get('imap_folder')

    get_bestellungen(imap_server, imap_user, imap_psw, imap_subject, imap_folder, pdffile)


if __name__ == "__main__": main()
