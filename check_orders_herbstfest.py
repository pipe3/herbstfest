#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# PROGRAMMER: oliver@kuhles.net
# DATE CREATED: 04/Oct/2021
# REVISED DATE:             <=(Date Revised - if any)
# PURPOSE: Herbstfest order summarize tool
#          To read formatted order emails from IMAP and summarize
#          them. Create a total order list as Excel. Create html stats
#          and upload them to our Wordpress site.
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

# Parse data from a data dict to get order details
def parse_tables(data):
    # initialize Bestellung b
    b = []

    # Zwiebel and Flamm as int
    b.append(int(data[3][1])) # Stueck Zwiebel
    b.append(int(data[4][1])) # Blech Zwiebel
    b.append(int(data[5][1])) # Blech Flamm
    b.append(int(data[6][1])) # Blech Flamm

    # Neuer Suesser
    if(data[7][1].startswith('keiner')):
        b.append(0)
        b.append(0)
        b.append(0)
    elif (data[7][1].startswith('1')):
        b.append(1)
        b.append(0)
        b.append(0)
    elif(data[7][1].startswith('2')):
        b.append(0)
        b.append(1)
        b.append(0)
    elif(data[7][1].startswith('3')):
        b.append(0)
        b.append(0)
        b.append(1)
    else:
        b.append('Fehler')
        b.append('Fehler')
        b.append('Fehler')

    # Vor or Fertig
    if(data[2][1].startswith('Vorgebacken')):
        b.append('Vor')
    elif (data[2][1].startswith('Fertig')):
        b.append('Fertig')
    else:
        b.append('Fehler')

    # Abholung
    if(data[9][1]).startswith('Abholung'):
        b.append('Abholung')
    elif(data[9][1]).startswith('Lieferung'):
        b.append('Lieferung')
    else:
        b.append('Fehler')

    # Tag
    if((data[10][1]).startswith('Samstag')):
        b.append('Samstag')
    elif((data[10][1]).startswith('Sonntag')):
        b.append('Sonntag')
    else:
        b.append('Fehler')

    b.append(data[11][1]) # Uhrzeit
    b.append(data[13][1]) # Vorname
    b.append(data[14][1]) # Name
    b.append(data[15][1]) # email
    b.append(data[16][1]) # Telefon

    # Addresse
    if(data[9][1] == 'Lieferung'):
        b.append(data[17][1])
    else:
        b.append('keine')

    return b


# Retrieve orders from IMAP and process them
def get_bestellungen(server, user, psw, subject, folder):

    with Imbox(server, user, psw,
        ssl=True,
        ssl_context=None,
        starttls=False) as imbox:

        # initialize vars
        bestellungen = []

        # Get all Herbstfest messages
        messages = imbox.messages(folder=folder, subject=subject)

        for uid, message in messages:

            # instantiate the parser and feed it
            p = HTMLTableParser()
            p.feed(message.body['plain'][0])

            # our data is sub-table 0
            data = p.tables[0]

            # get the bestellung as b
            b = parse_tables(data)

            # Append Bestellung to Bestellungen
            bestellungen.append(b)

    return bestellungen

# Create pandas dataframes from parsed orderes
def create_dataframes(bestellungen):
    # Define some headers
    cols = ['S Zwiebel', 'B Zwiebel', 'S Flamm', 'B Flamm',
            'Suesser 1L', 'Suesser 2L', 'Suesser 3L',
            'Art', 'Lieferung', 'Tag', 'Zeit',
            'Vorname', 'Name', 'email', 'Tel', 'Adresse']

    sum_list_stuecke_short = ['Sum Zwiebel Stuecke','Sum Flamm Stuecke', 'Sum Stuecke']
    sum_list_bleche_short = ['Sum Zwiebel Bleche','Sum Flamm Bleche', 'Sum Bleche']
    sum_list_suesser_short = ['Suesser 1L', 'Suesser 2L', 'Suesser 3L','Sum L Suesser']


    # Create empty dict to hold df's
    df_dict = {}

    ### All Bestellungen unsorted and enriched with sum columns
    b_df = pd.DataFrame(bestellungen, columns=cols)

    # Get the sum of Stuecke
    sum_stuecke = b_df['S Zwiebel'] + b_df['B Zwiebel'] * 9 + b_df['S Flamm'] + b_df['B Flamm'] * 9
    sum_zwiebel_stuecke = b_df['S Zwiebel'] + b_df['B Zwiebel'] * 9
    sum_flamm_stuecke = b_df['S Flamm'] + b_df['B Flamm'] * 9
    # Add the sum colums to the dataframe
    b_df['Sum Zwiebel Stuecke'] = sum_zwiebel_stuecke
    b_df['Sum Flamm Stuecke'] = sum_flamm_stuecke
    b_df['Sum Stuecke'] = sum_stuecke

    # Get the sum of Suesser
    sum_suesser = b_df['Suesser 1L'] + b_df['Suesser 2L'] * 2 + b_df['Suesser 3L'] * 3
    # Add the sum colums to the dataframe
    b_df['Sum L Suesser'] = sum_suesser

    # We dont need b_df in the dict, it serves as base for modifications below

    ### All Bestellungen sorted by Tag and Zeit
    sorted_df = b_df.sort_values(by=['Tag','Zeit']).reset_index()

    # Add it to the dict
    df_dict['Alle_Bestellungen'] = sorted_df

    ### Sum Bleche grouped by Tag and Zeit
    grouped_df = b_df.groupby(['Tag','Zeit','Art'])[sum_list_stuecke_short].sum().reset_index()

    # Get the sum of Bleche (round up)
    sum_zwiebel_bleche = grouped_df['Sum Zwiebel Stuecke'].div(9).apply(np.ceil).astype('int')
    sum_flamm_bleche = grouped_df['Sum Flamm Stuecke'].div(9).apply(np.ceil).astype('int')
    sum_bleche = sum_zwiebel_bleche + sum_flamm_bleche

    # Add the sum colums to the dataframe
    grouped_df['Sum Zwiebel Bleche'] = sum_zwiebel_bleche
    grouped_df['Sum Flamm Bleche'] = sum_flamm_bleche
    grouped_df['Sum Bleche'] = sum_bleche

    # Add it to the dict
    df_dict['Summe_Bleche_Zeit'] = grouped_df.drop(columns=sum_list_stuecke_short)


    ### Sum Bleche per Day
    day_df = b_df.groupby(['Tag','Art'])[sum_list_stuecke_short].sum().reset_index()

    # Get the sum of Bleche (round up)
    sum_zwiebel_bleche = day_df['Sum Zwiebel Stuecke'].div(9).apply(np.ceil).astype('int')
    sum_flamm_bleche = day_df['Sum Flamm Stuecke'].div(9).apply(np.ceil).astype('int')
    sum_bleche = sum_zwiebel_bleche + sum_flamm_bleche

    # Add the sum colums to the dataframe
    day_df['Sum Zwiebel Bleche'] = sum_zwiebel_bleche
    day_df['Sum Flamm Bleche'] = sum_flamm_bleche
    day_df['Sum Bleche'] = sum_bleche

    # Add it to the dict
    df_dict['Summe_Bleche_Tag'] = day_df.drop(columns=sum_list_stuecke_short)


    ### Total Sum Bleche
    bleche_total_sum_df = day_df.drop(columns=sum_list_stuecke_short+['Tag','Art']).sum().reset_index()

    # Add it to the dict
    df_dict['Total_Summe_Bleche'] = bleche_total_sum_df
    #df_dict['Total_Summe_Bleche'] = bleche_total_sum_df.drop([0,1], axis=0)


    ### Sum Suesser grouped by Tag and Zeit
    suesser_grouped_df = b_df.groupby(['Tag','Zeit'])[sum_list_suesser_short].sum().reset_index()
    # Add it to the dict
    df_dict['Summe_Suesser_Zeit'] = suesser_grouped_df

    ### Sum Suesser grouped by Tag
    suesser_day_df = b_df.groupby(['Tag'])[sum_list_suesser_short].sum().reset_index()
    # Add it to the dict
    df_dict['Summe_Suesser_Tag'] = suesser_day_df

    ## Total Sum Suesser
    suesser_total_sum_df = b_df[sum_list_suesser_short].sum().reset_index()
    #suesser_total_sum_df = suesser_day_df.sum().reset_index()
    df_dict['Total_Summe_Suesser'] = suesser_total_sum_df




    ## Alarm list sorted by max Bleche
    alarm_list = ['Tag','Zeit','Sum Bleche']
    alarm_df = grouped_df.sort_values(by=['Sum Bleche'], ascending=False)[alarm_list]
    # Add it to the dict
    df_dict['Alarm_Bleche'] = alarm_df

    # Return all dataframes as dict
    return df_dict


# Create Excel with tab for each dataframe within the dict
def write_to_excel(df_dict, filepath):
    with ExcelWriter(filepath) as writer:
        for key, val in df_dict.items():
            val.to_excel(writer, sheet_name=key)



def update_wordpress(df_dict, count_bestellungen, url, user, password, postid):

    credentials = user + ':' + password
    token = base64.b64encode(credentials.encode())
    header = {'Authorization': 'Basic ' + token.decode('utf-8')}

    # Create content
    content =  '<p>Letztes Update: '+str(datetime.now().strftime("%d-%m-%Y %H:%M:%S"))+'</p>'
    content += '<p>Anzahl Bestellungen: '+str(count_bestellungen)+'</p>'
    for key, val in df_dict.items():
        if key == 'Alle_Bestellungen':
            pass #dont print this to html because it contains sensitive data
        else:
            content += '<p><h4>'+key+'</h4>'
            content += '<figure class=\"wp-block-table dataframe\">'
            content += val.to_html()
            content += '</figure></p>'

    # create json for POST
    post = {
     'content'  : content
    }

    # post it to Wordpress
    response = requests.post(url + postid, headers=header, json=post)
    return response


def main():
    # Load the config file
    with open('config.json', 'r') as f:
        config = json.load(f)

    # Get vars from config file
    excelfile = config.get('excelfile')
    sleep_timer_min = config.get('sleep_timer_min')

    # Wordpress
    wp_url = config.get('wp_url')
    wp_user = config.get('wp_user')
    wp_pass = config.get('wp_pass')
    wp_postid = config.get('wp_postid')

    # IMAP
    imap_server = config.get('imap_server')
    imap_user = config.get('imap_user')
    imap_psw = config.get('imap_psw')
    imap_subject = config.get('imap_subject')
    imap_folder = config.get('imap_folder')


    # define vars
    b_count = 0

    print('{} Starte Herbstfest Bestellungen.'.format(datetime.now().strftime("%d-%m-%Y %H:%M:%S")))

    try:
        while True:
            print('{} Prüfe auf neue Bestellungen.'.format(datetime.now().strftime("%d-%m-%Y %H:%M:%S")))
            bestellungen = get_bestellungen(imap_server, imap_user, imap_psw, imap_subject, imap_folder)
            check_count = len(bestellungen)
            if check_count > b_count:
                start_time = time()
                print('Neue Bestellung gefunden. Anzahl Bestellungen jetzt: {}'.format(len(bestellungen)))
                df_dict = create_dataframes(bestellungen)

                write_to_excel(df_dict, excelfile)
                print('Excel file erstellt')

                wp_response = update_wordpress(df_dict, check_count, wp_url, wp_user, wp_pass, wp_postid)
                print('Wordpress Statistiken geupdated:', wp_response)
                print('Bestellungen geupdated in {:.3f}s.'.format(time() - start_time))

                b_count = check_count

            else:
                print('Keine neuen Bestellungen gefunden')

            # Pause until next run
            sleep(sleep_timer_min*60)

    except KeyboardInterrupt:
        print('{} Stoppe Herbstfest Bestellungen.'.format(datetime.now().strftime("%d-%m-%Y %H:%M:%S")))
        sys.exit()


if __name__ == "__main__": main()
