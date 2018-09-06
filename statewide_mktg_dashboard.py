import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from matplotlib import rcParams
from matplotlib.patches import Wedge, Circle
from matplotlib.collections import PatchCollection
import itertools
import sys
import win32com.client
import os

# figure out current working directory
cwd = os.getcwd()

# CALCULATIONS
# Pull in all the data we will need
inputs_csv = 'marketing-dashboard-inputs.csv'
inputs_df = pd.read_csv(inputs_csv)

# Some variables we need to set
rcParams['font.sans-serif'] = 'Arial'
rcParams['font.family'] = 'sans-serif'
rcParams['font.weight'] = 'bold'
rcParams['font.size'] = 8
rcParams['text.color'] = '#4b4b4b'
rcParams['axes.labelcolor'] = '#4b4b4b'
rcParams['xtick.color'] = '#4b4b4b'
rcParams['ytick.color'] = '#4b4b4b'
month_number = {1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
mo_num = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sept',10:'Oct',11:'Nov',12:'Dec'}
colliers_colors = (\
                   '#00467f',\
                   '#0093d0',\
                   '#4b4b4b',\
                   '#AAAAAA',\
                   '#E6E6E6',\
                   '#6CBFE6',\
                   '#DFEFF9',\
                   '#6583A7',\
                   '#4AB2DF',\
                   '#97999B',\
                   '#98AAC3',\
                   '#8AC6E7',\
                   '#D9DADC',\
                   '#CCD3E0',\
                   '#C1DEF1',\
                   '#FFC425',\
                   '#EE3124',\
                   '#B32317')
property_campaigns_list = [\
                           'Charleston Development Properties',\
                           'Charleston Flex Properties',\
                           'Charleston Industrial Properties',\
                           'Charleston Investment Properties'
                           'Charleston Land Properties',\
                           'Charleston Medical Properties',\
                           'Charleston Office Properties',\
                           'Charleston Retail Properties',\
                           'Columbia Development Properties',\
                           'Columbia Flex Properties',\
                           'Columbia Industrial Properties',\
                           'Columbia Investment Properties',\
                           'Columbia Land Properties',\
                           'Columbia Medical Properties',\
                           'Columbia Office Properties',\
                           'Columbia Retail Properties',\
                           'Greenville Development Properties',\
                           'Greenville Flex Properties',\
                           'Greenville Industrial Properties',\
                           'Greenville Investment Properties',\
                           'Greenville Land Properties',\
                           'Greenville Medical Properties',\
                           'Greenville Office Properties',\
                           'Greenville Retail Properties']
special_campaign_type = lambda c: c if (c in property_campaigns_list) or (c == 'Internal Newsletters 2018') else 'Custom Special Campaigns'
special_campaign_type2 = lambda c: c if (c in property_campaigns_list) or (c == 'Internal Newsletters 2018') or (c == 'Market Reports - 2018') else 'Custom Special Campaigns'
get_product_type = lambda x: x.split()[1]
turn_into_int = lambda x: int(x)
open_rate_calc = lambda row: round((row.Opened/row.Sent*100),1) if row.Sent > 1 else 0
click_rate_calc = lambda row: round((row.Clicked/row.Sent*100),2) if row.Sent > 1 else 0
round_click_rate = lambda x: round(x,1)
sent_perc_change = lambda row: int((row.Sent - row.Prev_Sent)/(row.Prev_Sent)*100) \
if row.Prev_Sent > 0 else 100
sent_perc_change_deg = lambda row: int((row.Sent - row.Prev_Sent)/(row.Prev_Sent)*360) \
if row.Prev_Sent > 0 else 360
def open_perc_change(row):
    try:
        return round((((row.Opened/row.Sent)-(row.Prev_Opened/row.Prev_Sent))*100),1)
    except ZeroDivisionError:
        if row.Prev_Sent == 0:
            return 100
        else:
            return -100
def click_perc_change(row):
    try:
        return round((((row.Clicked/row.Sent)-(row.Prev_Clicked/row.Prev_Sent))*100),2)
    except ZeroDivisionError:
        if row.Prev_Sent == 0:
            return 100
        else:
            return -100
def outer_theta1(p):
    if p >= 200:
        return 0
    elif p > 100:
        return 90 - int((p/100)*360)
    else:
        return 0
def outer_theta2(p):
    if p >= 200:
        return 360
    elif p > 100:
        return 90
    else:
        return 0
def outer_alpha(p):
    if p > 100:
        return 1
    else:
        return 0
def outline_alpha(p):
    if p > 100:
        return 1
    else:
        return 0 
def inner_theta1(p):
    if abs(p) >= 100:
        return 0
    if p == 0:
        return 0
    elif p > 0 & p < 100:
        return 90 - int((p/100)*360)
    elif p < 0 & p > -100:
        return 90  
def inner_theta2(p):
    if abs(p) >= 100:
        return 360
    if p == 0:
        return 0
    elif p > 0 & p < 100:
        return 90
    elif p < 0 & p > -100:
        return 90 + abs(int((p/100)*360))
def inner_alpha(p):
    if p == 0:
        return 0
    else:
        return 1
def patch_color(p):
    if p < 0:
        return '#B50900'
    else:
        return '#ffc425'

# Company-wide emails calculations
emails_csv = inputs_df.iloc[0]['CSV File Name or Numbers']
while '.csv' not in emails_csv:
    emails_csv = input('Please provide the name of the CSV file, including the .csv, for the two months of email data being compared: ')
emails_df = pd.read_csv(emails_csv)
emails_df = emails_df.rename(columns={'Send Date (Your time zone)':'Send Date'})
emails_df['Send Date'] = pd.to_datetime(emails_df['Send Date'])
emails_df['Send Month'] = emails_df['Send Date'].dt.month
if emails_df['Send Month'].max() == 12 & emails_df['Send Month'].min() == 1:
    rec_month = 1
    prev_month = 12
else:
    rec_month = emails_df['Send Month'].max()
    prev_month = emails_df['Send Month'].min()
recent_month_name = month_number[rec_month]
recent_mo = mo_num[rec_month]
prev_mo = mo_num[prev_month]
recent_emails = emails_df[emails_df['Send Month'] == rec_month]
prev_emails = emails_df[emails_df['Send Month'] == prev_month]
total_emails = recent_emails['Sent'].sum()
total_opens = recent_emails['Opened'].sum()
total_open_rate = round(((total_opens/total_emails)*100),2)
total_clicks = recent_emails['Clicked'].sum()
total_click_rate = round(((total_clicks/total_emails)*100),2)
total_opens = '{:,}'.format(total_opens)
total_clicks = '{:,}'.format(total_clicks)

# Email recipients by campaign
combined_emails = recent_emails[['Campaign','Sent']].reset_index(drop=True)
combined_emails['Long_Email_Campaign'] = combined_emails.Campaign.apply(special_campaign_type2)
combined_emails['Email_Campaign'] = combined_emails.Long_Email_Campaign.apply(get_product_type)
grouped_combined_emails = combined_emails.groupby('Email_Campaign').Sent.sum().reset_index()
total_emails_sent = grouped_combined_emails.Sent.sum()
grouped_combined_emails['Perc_Total'] = grouped_combined_emails.Sent.apply(lambda s: int((s/total_emails_sent)*100))
grouped_combined_emails['Label'] = grouped_combined_emails.apply(lambda row: str(row.Perc_Total) + "%\n" + row.Email_Campaign, axis=1)
grouped_combined_emails = grouped_combined_emails.sort_values(by=['Email_Campaign'],ascending=True).reset_index(drop=True)
perc_count = grouped_combined_emails['Perc_Total'].tolist()
campaign_labels = grouped_combined_emails['Label'].tolist()
fig9, ax9 = plt.subplots(figsize=(4, 4), subplot_kw=dict(aspect="equal"))
wedges9,texts9 = ax9.pie(perc_count,
                        colors=colliers_colors,
                        labels=campaign_labels,
                        labeldistance=1.2,
                        pctdistance=0.85,
                        startangle=90)
for w in wedges9:
    w.set_width(0.25)
    w.set_linewidth(1)
    w.set_edgecolor('white')
for t in texts9:
    t.set_horizontalalignment('center')

total_prev_emails = prev_emails['Sent'].sum()
sent_perc = int((total_emails - total_prev_emails)/(total_prev_emails)*100)
ax9.add_artist(Wedge((0,0), 0.7, inner_theta1(sent_perc), inner_theta2(sent_perc), color=patch_color(sent_perc), alpha=inner_alpha(sent_perc)))
ax9.add_artist(Circle((0,0),0.55,color="#ffffff",alpha=1))
ax9.add_artist(Circle((0,0),0.525,color="#00467f",alpha=1))
plt.tight_layout()
fig9.savefig('total_emails_by_campaign.eps',transparent=True)

total_sent_emails = '{:,}'.format(total_emails)
total_prev_sent_emails = '{:,}'.format(total_prev_emails)

# Top performing emails
recent_property_emails = recent_emails[['Email Name','Campaign','Sent','Open Rate','Click Rate','Unsubscribed','Hard Bounced','Soft Bounced']].reset_index(drop=True)
recent_property_emails = recent_property_emails[recent_property_emails.Campaign.isin(property_campaigns_list)]
recent_property_emails = recent_property_emails.sort_values(by=['Click Rate'],ascending=False).reset_index(drop=True)
top_property_emails = recent_property_emails.iloc[:5].reset_index()
top_property_emails['Product'] = top_property_emails.Campaign.apply(get_product_type)
top_property_emails['Open_Rate'] = top_property_emails['Open Rate'].apply(turn_into_int)
top_property_emails['Click_Rate'] = top_property_emails['Click Rate'].apply(round_click_rate)
top_property_emails['Bounce'] = top_property_emails['Hard Bounced'] + top_property_emails['Soft Bounced']
top_property_emails['Sent'] = top_property_emails['Sent'].apply(lambda x: '{:,}'.format(x))
top_property_emails['Open_Rate'] = top_property_emails['Open_Rate'].apply(lambda x: str(x) + '%')
top_property_emails['Click_Rate'] = top_property_emails['Click_Rate'].apply(lambda x: str(x) + '%')
top_property_emails['Bounce'] = top_property_emails['Bounce'].apply(lambda x: str(x))
top_property_emails['Unsubscribed'] = top_property_emails['Unsubscribed'].apply(lambda x: str(x))
top_property_emails = top_property_emails.iloc[:5].reset_index()
top_emails = top_property_emails[['Email Name','Product','Sent','Open_Rate','Click_Rate','Unsubscribed','Bounce']]
top_emails_values_list = top_emails.values.tolist()
top_emails_values = [y for x in top_emails_values_list for y in x]
top_emails_columns_list = ['Email Name','Product','Sent','Open Rate','Click Rate','Unsubscribed','Bounces']
top_emails_list = top_emails_columns_list + top_emails_values

# Email change month over month by campaign
recent_campaigns = recent_emails[['Campaign','Sent','Opened','Clicked']].reset_index(drop=True)
recent_campaigns = recent_campaigns[recent_campaigns.Campaign.isin(property_campaigns_list)]
recent_campaigns_group = recent_campaigns.groupby('Campaign')['Sent','Opened','Clicked'].sum().reset_index()
prev_campaigns = prev_emails[['Campaign','Sent','Opened','Clicked']].reset_index(drop=True)
prev_campaigns.rename(columns={'Sent':'Prev_Sent','Opened':'Prev_Opened','Clicked':'Prev_Clicked'},inplace=True)
prev_campaigns = prev_campaigns[prev_campaigns.Campaign.isin(property_campaigns_list)]
prev_campaigns_group = prev_campaigns.groupby('Campaign')['Prev_Sent','Prev_Opened','Prev_Clicked'].sum().reset_index()
campaigns_comparison = pd.merge(recent_campaigns_group,prev_campaigns_group,how='outer')
campaigns_comparison['Product'] = campaigns_comparison.Campaign.apply(get_product_type)
campaign_totals = campaigns_comparison.groupby('Product')['Sent','Prev_Sent','Opened','Prev_Opened','Clicked','Prev_Clicked'].sum().reset_index()
campaign_totals['Open_Rate'] = campaign_totals.apply(open_rate_calc,axis=1)
campaign_totals['Click_Rate'] = campaign_totals.apply(click_rate_calc,axis=1)
campaign_totals['Sent_Change'] = campaign_totals.apply(sent_perc_change,axis=1)
campaign_totals['Open_Change'] = campaign_totals.apply(open_perc_change,axis=1)
campaign_totals['Click_Change'] = campaign_totals.apply(click_perc_change,axis=1)
campaign_totals['Total_Emails'] = campaign_totals.Sent.apply(turn_into_int)
campaign_totals['Prev_Total'] = campaign_totals.Prev_Sent.apply(turn_into_int)
campaign_totals_display = campaign_totals[['Product','Total_Emails','Sent_Change','Prev_Total','Open_Rate','Open_Change','Click_Rate','Click_Change']]

# DEVELOPMENT
try:
    development_totals = campaign_totals_display[campaign_totals_display.Product == 'Development'].reset_index(drop=True)
    dev_emails = development_totals.iloc[0]['Total_Emails']
    dev_perc = development_totals.iloc[0]['Sent_Change']
    dev_prev_emails = development_totals.iloc[0]['Prev_Total']
    dev_open_rate = development_totals.iloc[0]['Open_Rate']
    dev_open_change = development_totals.iloc[0]['Open_Change']
    dev_click_rate = development_totals.iloc[0]['Click_Rate']
    dev_click_change = development_totals.iloc[0]['Click_Change']
except IndexError:
    dev_emails = 0
    dev_perc = 0
    dev_prev_emails = 0
    dev_open_rate = 0
    dev_open_change = 0
    dev_click_rate = 0
    dev_click_change = 0
fig1=plt.figure(figsize=(2.3,2.3))
ax1=fig1.add_subplot(111,aspect='equal')
ax1.add_artist(Wedge((.5,.5), 0.5, outer_theta1(dev_perc), outer_theta2(dev_perc), color=patch_color(dev_perc), alpha=outer_alpha(dev_perc)))
ax1.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(dev_perc)))
ax1.add_artist(Wedge((.5,.5), 0.425, inner_theta1(dev_perc), inner_theta2(dev_perc), color=patch_color(dev_perc), alpha=inner_alpha(dev_perc)))
ax1.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax1.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig1.savefig('development-emails.eps',transparent=True)
dev_emails = '{:,}'.format(dev_emails)
dev_prev_emails = '{:,}'.format(dev_prev_emails)

# FLEX
try:
    flex_totals = campaign_totals_display[campaign_totals_display.Product == 'Flex'].reset_index(drop=True)
    flex_emails = flex_totals.iloc[0]['Total_Emails']
    flex_perc = flex_totals.iloc[0]['Sent_Change']
    flex_prev_emails = flex_totals.iloc[0]['Prev_Total']
    flex_open_rate = flex_totals.iloc[0]['Open_Rate']
    flex_open_change = flex_totals.iloc[0]['Open_Change']
    flex_click_rate = flex_totals.iloc[0]['Click_Rate']
    flex_click_change = flex_totals.iloc[0]['Click_Change']
except IndexError:
    flex_emails = 0
    flex_perc = 0
    flex_prev_emails = 0
    flex_open_rate = 0
    flex_open_change = 0
    flex_click_rate = 0
    flex_click_change = 0
fig2=plt.figure(figsize=(2.3,2.3))
ax2=fig2.add_subplot(111,aspect='equal')
ax2.add_artist(Wedge((.5,.5), 0.5, outer_theta1(flex_perc), outer_theta2(flex_perc), color=patch_color(flex_perc), alpha=outer_alpha(flex_perc)))
ax2.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(flex_perc)))
ax2.add_artist(Wedge((.5,.5), 0.425, inner_theta1(flex_perc), inner_theta2(flex_perc), color=patch_color(flex_perc), alpha=inner_alpha(flex_perc)))
ax2.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax2.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig2.savefig('flex-emails.eps',transparent=True)
flex_emails = '{:,}'.format(flex_emails)
flex_prev_emails = '{:,}'.format(flex_prev_emails)

# INDUSTRIAL
try:
    industrial_totals = campaign_totals_display[campaign_totals_display.Product == 'Industrial'].reset_index(drop=True)
    ind_emails = industrial_totals.iloc[0]['Total_Emails']
    ind_perc = industrial_totals.iloc[0]['Sent_Change']
    ind_prev_emails = industrial_totals.iloc[0]['Prev_Total']
    ind_open_rate = industrial_totals.iloc[0]['Open_Rate']
    ind_open_change = industrial_totals.iloc[0]['Open_Change']
    ind_click_rate = industrial_totals.iloc[0]['Click_Rate']
    ind_click_change = industrial_totals.iloc[0]['Click_Change']
except IndexError:
    ind_emails = 0
    ind_perc = 0
    ind_prev_emails = 0
    ind_open_rate = 0
    ind_open_change = 0
    ind_click_rate = 0
    ind_click_change = 0
fig3=plt.figure(figsize=(2.3,2.3))
ax3=fig3.add_subplot(111,aspect='equal')
ax3.add_artist(Wedge((.5,.5), 0.5, outer_theta1(ind_perc), outer_theta2(ind_perc), color=patch_color(ind_perc), alpha=outer_alpha(ind_perc)))
ax3.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(ind_perc)))
ax3.add_artist(Wedge((.5,.5), 0.425, inner_theta1(ind_perc), inner_theta2(ind_perc), color=patch_color(ind_perc), alpha=inner_alpha(ind_perc)))
ax3.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax3.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig3.savefig('industrial-emails.eps',transparent=True)
ind_emails = '{:,}'.format(ind_emails)
ind_prev_emails = '{:,}'.format(ind_prev_emails)

# INVESTMENT
try:
    investment_totals = campaign_totals_display[campaign_totals_display.Product == 'Investment'].reset_index(drop=True)
    inv_emails = investment_totals.iloc[0]['Total_Emails']
    inv_perc = investment_totals.iloc[0]['Sent_Change']
    inv_prev_emails = investment_totals.iloc[0]['Prev_Total']
    inv_open_rate = investment_totals.iloc[0]['Open_Rate']
    inv_open_change = investment_totals.iloc[0]['Open_Change']
    inv_click_rate = investment_totals.iloc[0]['Click_Rate']
    inv_click_change = investment_totals.iloc[0]['Click_Change']
except IndexError:
    inv_emails = 0
    inv_perc = 0
    inv_prev_emails = 0
    inv_open_rate = 0
    inv_open_change = 0
    inv_click_rate = 0
    inv_click_change = 0
fig4=plt.figure(figsize=(2.3,2.3))
ax4=fig4.add_subplot(111,aspect='equal')
ax4.add_artist(Wedge((.5,.5), 0.5, outer_theta1(inv_perc), outer_theta2(inv_perc), color=patch_color(inv_perc), alpha=outer_alpha(inv_perc)))
ax4.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(inv_perc)))
ax4.add_artist(Wedge((.5,.5), 0.425, inner_theta1(inv_perc), inner_theta2(inv_perc), color=patch_color(inv_perc), alpha=inner_alpha(inv_perc)))
ax4.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax4.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig4.savefig('investment-emails.eps',transparent=True)
inv_emails = '{:,}'.format(inv_emails)
inv_prev_emails = '{:,}'.format(inv_prev_emails)

# LAND
try:
    land_totals = campaign_totals_display[campaign_totals_display.Product == 'Land'].reset_index(drop=True)
    land_emails = land_totals.iloc[0]['Total_Emails']
    land_perc = land_totals.iloc[0]['Sent_Change']
    land_prev_emails = land_totals.iloc[0]['Prev_Total']
    land_open_rate = land_totals.iloc[0]['Open_Rate']
    land_open_change = land_totals.iloc[0]['Open_Change']
    land_click_rate = land_totals.iloc[0]['Click_Rate']
    land_click_change = land_totals.iloc[0]['Click_Change']
except IndexError:
    land_emails = 0
    land_perc = 0
    land_prev_emails = 0
    land_open_rate = 0
    land_open_change = 0
    land_click_rate = 0
    land_click_change = 0
fig5=plt.figure(figsize=(2.3,2.3))
ax5=fig5.add_subplot(111,aspect='equal')
ax5.add_artist(Wedge((.5,.5), 0.5, outer_theta1(land_perc), outer_theta2(land_perc), color=patch_color(land_perc), alpha=outer_alpha(land_perc)))
ax5.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(land_perc)))
ax5.add_artist(Wedge((.5,.5), 0.425, inner_theta1(land_perc), inner_theta2(land_perc), color=patch_color(land_perc), alpha=inner_alpha(land_perc)))
ax5.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax5.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig5.savefig('land-emails.eps',transparent=True)
land_emails = '{:,}'.format(land_emails)
land_prev_emails = '{:,}'.format(land_prev_emails)

# MEDICAL
try:
    medical_totals = campaign_totals_display[campaign_totals_display.Product == 'Medical'].reset_index(drop=True)
    med_emails = medical_totals.iloc[0]['Total_Emails']
    med_perc = medical_totals.iloc[0]['Sent_Change']
    med_prev_emails = medical_totals.iloc[0]['Prev_Total']
    med_open_rate = medical_totals.iloc[0]['Open_Rate']
    med_open_change = medical_totals.iloc[0]['Open_Change']
    med_click_rate = medical_totals.iloc[0]['Click_Rate']
    med_click_change = medical_totals.iloc[0]['Click_Change']
except IndexError:
    med_emails = 0
    med_perc = 0
    med_prev_emails = 0
    med_open_rate = 0
    med_open_change = 0
    med_click_rate = 0
    med_click_change = 0
fig6=plt.figure(figsize=(2.3,2.3))
ax6=fig6.add_subplot(111,aspect='equal')
ax6.add_artist(Wedge((.5,.5), 0.5, outer_theta1(med_perc), outer_theta2(med_perc), color=patch_color(med_perc), alpha=outer_alpha(med_perc)))
ax6.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(med_perc)))
ax6.add_artist(Wedge((.5,.5), 0.425, inner_theta1(med_perc), inner_theta2(med_perc), color=patch_color(med_perc), alpha=inner_alpha(med_perc)))
ax6.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax6.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig6.savefig('medical-emails.eps',transparent=True)
med_emails = '{:,}'.format(med_emails)
med_prev_emails = '{:,}'.format(med_prev_emails)

# OFFICE
try:
    office_totals = campaign_totals_display[campaign_totals_display.Product == 'Office'].reset_index(drop=True)
    off_emails = office_totals.iloc[0]['Total_Emails']
    off_perc = office_totals.iloc[0]['Sent_Change']
    off_prev_emails = office_totals.iloc[0]['Prev_Total']
    off_open_rate = office_totals.iloc[0]['Open_Rate']
    off_open_change = office_totals.iloc[0]['Open_Change']
    off_click_rate = office_totals.iloc[0]['Click_Rate']
    off_click_change = office_totals.iloc[0]['Click_Change']
except IndexError:
    off_emails = 0
    off_perc = 0
    off_prev_emails = 0
    off_open_rate = 0
    off_open_change = 0
    off_click_rate = 0
    off_click_change = 0
fig7=plt.figure(figsize=(2.3,2.3))
ax7=fig7.add_subplot(111,aspect='equal')
ax7.add_artist(Wedge((.5,.5), 0.5, outer_theta1(off_perc), outer_theta2(off_perc), color=patch_color(off_perc), alpha=outer_alpha(off_perc)))
ax7.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(off_perc)))
ax7.add_artist(Wedge((.5,.5), 0.425, inner_theta1(off_perc), inner_theta2(off_perc), color=patch_color(off_perc), alpha=inner_alpha(off_perc)))
ax7.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax7.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig7.savefig('office-emails.eps',transparent=True)
off_emails = '{:,}'.format(off_emails)
off_prev_emails = '{:,}'.format(off_prev_emails)

# RETAIL
try:
    retail_totals = campaign_totals_display[campaign_totals_display.Product == 'Retail'].reset_index(drop=True)
    ret_emails = retail_totals.iloc[0]['Total_Emails']
    ret_perc = retail_totals.iloc[0]['Sent_Change']
    ret_prev_emails = retail_totals.iloc[0]['Prev_Total']
    ret_open_rate = retail_totals.iloc[0]['Open_Rate']
    ret_open_change = retail_totals.iloc[0]['Open_Change']
    ret_click_rate = retail_totals.iloc[0]['Click_Rate']
    ret_click_change = retail_totals.iloc[0]['Click_Change']
except IndexError:
    ret_emails = 0
    ret_perc = 0
    ret_prev_emails = 0
    ret_open_rate = 0
    ret_open_change = 0
    ret_click_rate = 0
    ret_click_change = 0
fig8=plt.figure(figsize=(2.3,2.3))
ax8=fig8.add_subplot(111,aspect='equal')
ax8.add_artist(Wedge((.5,.5), 0.5, outer_theta1(ret_perc), outer_theta2(ret_perc), color=patch_color(ret_perc), alpha=outer_alpha(ret_perc)))
ax8.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(ret_perc)))
ax8.add_artist(Wedge((.5,.5), 0.425, inner_theta1(ret_perc), inner_theta2(ret_perc), color=patch_color(ret_perc), alpha=inner_alpha(ret_perc)))
ax8.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax8.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig8.savefig('retail-emails.eps',transparent=True)
ret_emails = '{:,}'.format(ret_emails)
ret_prev_emails = '{:,}'.format(ret_prev_emails)

# Social Media Impressions
facebook_csv = inputs_df.iloc[1]['CSV File Name or Numbers']
while '.csv' not in facebook_csv:
    facebook_csv = input('Please provide the name of the CSV file, including the .csv, for the two months of Facebook post data being compared: ')
rec_twitter_csv = inputs_df.iloc[2]['CSV File Name or Numbers']
while '.csv' not in rec_twitter_csv:
    rec_twitter_csv = input('Please provide the name of the CSV file, including the .csv, for the **most recent** month of Twitter tweet data being compared: ')
prev_twitter_csv = inputs_df.iloc[3]['CSV File Name or Numbers']
while '.csv' not in prev_twitter_csv:
    prev_twitter_csv = input('Please provide the name of the CSV file, including the .csv, for the **previous** month of Twitter tweet data being compared: ')
instagram_csv = inputs_df.iloc[4]['CSV File Name or Numbers']
while '.csv' not in instagram_csv:
    instagram_csv = input('Please provide the name of the CSV file, including the .csv, for the two months of Instagram post data being compared: ')
facebook_posts = pd.read_csv(facebook_csv)
facebook_posts = facebook_posts.drop(facebook_posts.index[0]).reset_index(drop=True)
facebook_posts = facebook_posts.rename(columns={'Post Message':'Post_Nickname','Lifetime Post Total Impressions':'Impressions','Lifetime Post Audience Targeting Unique Consumptions by Type - link clicks':'Clicks','Lifetime Matched Audience Targeting Consumptions on Post':'Interactions'})
facebook_posts['Posted'] = pd.to_datetime(facebook_posts['Posted'])
facebook_posts['Posted Month'] = facebook_posts['Posted'].dt.month
recent_facebook_posts = facebook_posts[facebook_posts['Posted Month'] == rec_month]
prev_facebook_posts = facebook_posts[facebook_posts['Posted Month'] == prev_month]
recent_facebook_data = recent_facebook_posts[['Post_Nickname','Impressions','Clicks','Interactions']]
recent_facebook_data['Source'] = 'Facebook'
recent_facebook_data = recent_facebook_data.fillna(value=0).reset_index(drop=True)
prev_facebook_data = prev_facebook_posts[['Post_Nickname','Impressions','Clicks','Interactions']]
prev_facebook_data['Source'] = 'Facebook'
prev_facebook_data = prev_facebook_data.fillna(value=0).reset_index(drop=True)

recent_twitter_posts = pd.read_csv(rec_twitter_csv)
recent_twitter_data = recent_twitter_posts[['Tweet text','impressions','url clicks','engagements']]
recent_twitter_data = recent_twitter_data.rename(columns={'Tweet text':'Post_Nickname','impressions':'Impressions','url clicks':'Clicks','engagements':'Interactions'})
recent_twitter_data['Source'] = 'Twitter'
prev_twitter_posts = pd.read_csv(prev_twitter_csv)
prev_twitter_data = prev_twitter_posts[['Tweet text','impressions','url clicks','engagements']]
prev_twitter_data = prev_twitter_data.rename(columns={'Tweet text':'Post_Nickname','impressions':'Impressions','url clicks':'Clicks','engagements':'Interactions'})
prev_twitter_data['Source'] = 'Twitter'

insta_posts = pd.read_csv(instagram_csv)
insta_posts['Clicks'] = 0
insta_posts['Interactions'] = insta_posts['Likes'] + insta_posts['Comments']
insta_posts['Source'] = 'Instagram'
insta_posts = insta_posts.rename(columns={'Reach':'Impressions'})
insta_posts['Post_Date'] = pd.to_datetime(insta_posts['Post_Date'])
insta_posts['Posted Month'] = insta_posts['Post_Date'].dt.month
recent_insta_posts = insta_posts[insta_posts['Posted Month'] == rec_month]
prev_insta_posts = insta_posts[insta_posts['Posted Month'] == prev_month]
recent_insta_data = recent_insta_posts[['Post_Nickname','Impressions','Clicks','Interactions','Source']]

recent_posts_data = pd.concat([recent_facebook_data,recent_twitter_data,recent_insta_data])
prev_posts_data = pd.concat([prev_facebook_data,prev_twitter_data])
recent_posts_data['Impressions'] = pd.to_numeric(recent_posts_data['Impressions'],downcast='integer').reset_index(drop=True)
recent_posts_data['Clicks'] = pd.to_numeric(recent_posts_data['Clicks'],downcast='integer').reset_index(drop=True)
recent_posts_data['Interactions'] = pd.to_numeric(recent_posts_data['Interactions'],downcast='integer').reset_index(drop=True)
recent_impressions_total = recent_posts_data['Impressions'].sum()
prev_posts_data['Impressions'] = pd.to_numeric(prev_posts_data['Impressions'],downcast='integer').reset_index(drop=True)
prev_impressions_total = prev_posts_data['Impressions'].sum()
impressions_perc_change = int(((recent_impressions_total-prev_impressions_total)/prev_impressions_total*100))
fig11=plt.figure(figsize=(2.3,2.3))
ax11=fig11.add_subplot(111,aspect='equal')
ax11.add_artist(Wedge((.5,.5), 0.5, outer_theta1(impressions_perc_change), outer_theta2(impressions_perc_change), color=patch_color(impressions_perc_change), alpha=outer_alpha(impressions_perc_change)))
ax11.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(impressions_perc_change)))
ax11.add_artist(Wedge((.5,.5), 0.425, inner_theta1(impressions_perc_change), inner_theta2(impressions_perc_change), color=patch_color(impressions_perc_change), alpha=inner_alpha(impressions_perc_change)))
ax11.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax11.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig11.savefig('social-impressions.eps',transparent=True)
recent_impressions_total = '{:,}'.format(recent_impressions_total)
prev_impressions_total = '{:,}'.format(prev_impressions_total)

# Social Media Network Followers & Posts
count_twitter_followers = int(inputs_df.iloc[5]['CSV File Name or Numbers'])
count_twitter_posts = recent_twitter_data['Post_Nickname'].count()
count_fb_followers = int(inputs_df.iloc[6]['CSV File Name or Numbers'])
count_fb_posts = recent_facebook_data['Post_Nickname'].count()
count_insta_followers = int(inputs_df.iloc[7]['CSV File Name or Numbers'])
count_insta_posts = recent_insta_data['Post_Nickname'].count()
count_twitter_followers = '{:,}'.format(count_twitter_followers)
count_twitter_posts = '{:,}'.format(count_twitter_posts)
count_fb_followers = '{:,}'.format(count_fb_followers)
count_fb_posts = '{:,}'.format(count_fb_posts)
count_insta_followers = '{:,}'.format(count_insta_followers)
count_insta_posts = '{:,}'.format(count_insta_posts)

# Top Performing Social Posts
recent_posts_data = recent_posts_data.sort_values(by=['Clicks'],ascending=False).reset_index(drop=True)
top_social_posts = recent_posts_data.iloc[:5]
top_social_posts = top_social_posts.rename(columns={'Post_Nickname':'Social Post','Source':'Network'})
top_social_posts = top_social_posts[['Social Post','Network','Impressions','Clicks','Interactions']]
top_social_posts = top_social_posts.replace(regex=True, to_replace="'", value="")
top_social_posts = top_social_posts.replace(regex=True, to_replace="\n", value="")
top_social_posts['Impressions'] = top_social_posts['Impressions'].apply(lambda x: '{:,}'.format(x))
top_social_posts['Clicks'] = top_social_posts['Clicks'].apply(lambda x: '{:,}'.format(x))
top_social_posts['Interactions'] = top_social_posts['Interactions'].apply(lambda x: '{:,}'.format(x))
top_social_posts_values_list = top_social_posts.values.tolist()
top_social_posts_values = [y for x in top_social_posts_values_list for y in x]
top_social_posts_columns_list = top_social_posts.columns.tolist()
top_social_posts_list = top_social_posts_columns_list + top_social_posts_values

# Website Visits by Source
website_page_views_csv = inputs_df.iloc[8]['CSV File Name or Numbers']
while '.csv' not in website_page_views_csv:
    website_page_views_csv = input('Please provide the name of the CSV file, including the .csv, for the month of website data: ')
page_views_df = pd.read_csv(website_page_views_csv)
page_views_df = page_views_df.drop(page_views_df.index[-1]).reset_index(drop=True)
page_views_df = page_views_df[page_views_df.Name.str.contains('localhost') == False]
page_views_df['Domain'] = page_views_df.Name.apply(lambda x: x.split('.')[1])
colliers_page_views = page_views_df[page_views_df.Domain == 'colliers']
colliers_page_views_totals = colliers_page_views.groupby('Domain').sum()
colliers_page_views_totals['Total Visits'] = colliers_page_views_totals.apply(lambda row: int(sum(row[1:])),axis=1)
total_page_views = int(colliers_page_views_totals.iloc[0]['Total Visits'])
total_page_views = '{:,}'.format(total_page_views)

website_sources_csv = inputs_df.iloc[9]['CSV File Name or Numbers']
while '.csv' not in website_sources_csv:
    website_sources_csv = input('Please provide the name of the CSV file, including the .csv, for the month of website data: ')
website_sources_df = pd.read_csv(website_sources_csv)
website_sources_df = website_sources_df.drop(website_sources_df.index[-3:]).reset_index(drop=True)
website_sources_df = website_sources_df[['Name','Sessions']]
website_sources_df = website_sources_df[website_sources_df.Name != 'Offline']
website_sources_df = website_sources_df[website_sources_df.Name != 'Other Campaigns']
website_sources_df['Name'] = website_sources_df['Name'].apply(lambda n: n.split()[0])
website_sources_df = website_sources_df.sort_values(by=['Name'],ascending=True).reset_index(drop=True)
total_sessions = website_sources_df.Sessions.sum()
website_sources_df['Percentage'] = website_sources_df['Sessions'].apply(lambda x: str(int(100*x/total_sessions))).reset_index(drop=True)
website_sources_df['Label'] = website_sources_df.apply(lambda row: row.Percentage + "%\n" + row.Name, axis=1)
sessions_count = website_sources_df['Sessions'].tolist()
channel_labels = website_sources_df['Label'].tolist()
fig10, ax10 = plt.subplots(figsize=(3.1,3.1), subplot_kw=dict(aspect="equal"))
wedges10,texts10 = ax10.pie(sessions_count,
                        colors=colliers_colors,
                        labels=channel_labels,
                        labeldistance=1.2,
                        pctdistance=0.85,
                        startangle=90)
for w in wedges10:
    w.set_width(0.25)
    w.set_linewidth(2)
    w.set_edgecolor('white')
for t in texts10:
    t.set_horizontalalignment('center')
ax10.add_artist(Circle((0,0),0.7,color="#00467f",alpha=1))
plt.tight_layout()
fig10.savefig('website-visits.eps',transparent=True)

# Website Visits by Campaign
sessions_by_campaigns_csv = inputs_df.iloc[10]['CSV File Name or Numbers']
while '.csv' not in sessions_by_campaigns_csv:
    sessions_by_campaigns_csv = input('Please provide the name of the CSV file, including the .csv, for the month of sessions data: ')
campaign_sessions_df = pd.read_csv(sessions_by_campaigns_csv)
campaign_sessions_df = campaign_sessions_df.drop(campaign_sessions_df.index[-1]).reset_index(drop=True)
campaign_sessions_df = campaign_sessions_df.sort_values(by=['Name'],ascending=True).reset_index(drop=True)
campaign_sessions_df['Total Visits'] = campaign_sessions_df.apply(lambda row: int(sum(row[1:])),axis=1)
campaign_visits = campaign_sessions_df[['Name','Total Visits']].reset_index(drop=True)
campaign_visits = campaign_visits.replace(regex=True, to_replace="Columbia ", value="").reset_index(drop=True)
campaign_visits = campaign_visits.replace(regex=True, to_replace="Greenville ", value="").reset_index(drop=True)
campaign_visits = campaign_visits.replace(regex=True, to_replace="Charleston ", value="").reset_index(drop=True)
campaign_visits_group = campaign_visits.groupby('Name')['Total Visits'].sum().reset_index()
campaign_visits_group = campaign_visits_group[campaign_visits_group.Name != 'REMS Property Newsletters']
campaign_visits_group = campaign_visits_group[campaign_visits_group.Name != 'Internal Newsletters 2018']
campaign_visits_group = campaign_visits_group[campaign_visits_group['Total Visits'] != 0]
campaign_visits_group = campaign_visits_group.sort_values(by=['Total Visits'],ascending=False).reset_index(drop=True)
campaign_visits_group['Total Visits'] = campaign_visits_group['Total Visits'].apply(lambda x: '{:,}'.format(x))
campaign_visits_values_list = campaign_visits_group.values.tolist()
campaign_visits_values = [y for x in campaign_visits_values_list for y in x]
campaign_visits_columns_list = campaign_visits_group.columns.tolist()
campaign_visits_list = campaign_visits_columns_list + campaign_visits_values

# Website Visits by Page Type
colliers_page_views = colliers_page_views.replace(regex=True, to_replace="/en", value="").reset_index(drop=True)
colliers_page_views = colliers_page_views.replace(regex=True, to_replace="/united-states", value="").reset_index(drop=True)
colliers_page_views['Page Type'] = colliers_page_views['Name'].apply(lambda p: p.split('/')[1])
colliers_page_views['Page Type'] = colliers_page_views['Page Type'].apply(lambda p: p.capitalize())
colliers_page_groups = colliers_page_views.groupby('Page Type').sum()
colliers_page_groups['Total Visits'] = colliers_page_groups.apply(lambda row: int(sum(row)),axis=1)
page_views_table = colliers_page_groups[['Total Visits']]
page_views_table.reset_index(level=0,inplace=True)
page_views_table = page_views_table.sort_values(by=['Total Visits'],ascending=False).reset_index(drop=True)
page_views_table['Total Visits'] = page_views_table['Total Visits'].apply(lambda x: '{:,}'.format(x))
page_views_table_values_list = page_views_table.values.tolist()
page_views_table_values = [y for x in page_views_table_values_list for y in x]
page_views_table_columns_list = page_views_table.columns.tolist()
page_views_table_list = page_views_table_columns_list + page_views_table_values

# Subscriber Leads
current_prop_subscribers = int(inputs_df.iloc[11]['CSV File Name or Numbers'])
current_prop_subscribers = '{:,}'.format(current_prop_subscribers)
current_research_subscribers = int(inputs_df.iloc[12]['CSV File Name or Numbers'])
current_research_subscribers = '{:,}'.format(current_research_subscribers)

# Press Releases
ytd_prs = int(inputs_df.iloc[13]['CSV File Name or Numbers'])
prev_ytd_prs = int(inputs_df.iloc[14]['CSV File Name or Numbers'])
prs_perc_change = int(((ytd_prs-prev_ytd_prs)/prev_ytd_prs*100))

fig12=plt.figure(figsize=(2.3,2.3))
ax12=fig12.add_subplot(111,aspect='equal')
ax12.add_artist(Wedge((.5,.5), 0.5, outer_theta1(prs_perc_change), outer_theta2(prs_perc_change), color=patch_color(prs_perc_change), alpha=outer_alpha(prs_perc_change)))
ax12.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(prs_perc_change)))
ax12.add_artist(Wedge((.5,.5), 0.425, inner_theta1(prs_perc_change), inner_theta2(prs_perc_change), color=patch_color(prs_perc_change), alpha=inner_alpha(prs_perc_change)))
ax12.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax12.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig12.savefig('press-releases.eps',transparent=True)

# Market Reports
qtrly_emails_csv = inputs_df.iloc[15]['CSV File Name or Numbers']
while '.csv' not in qtrly_emails_csv:
    qtrly_emails_csv = input('Please provide the name of the CSV file, including the .csv, for the email data being compared: ')
mr_emails_df = pd.read_csv(qtrly_emails_csv)
mr_emails_df = mr_emails_df[mr_emails_df['Campaign'].str.contains('Market Reports')==True].reset_index(drop=True)
mr_emails_df['Year'] = mr_emails_df['Email Name'].apply(lambda x: int(x.split()[0]))
mr_emails_df['Quarter'] = mr_emails_df['Email Name'].apply(lambda x: x.split()[1])
mr_emails_group = mr_emails_df.groupby(['Year','Quarter'])['Sent'].sum().reset_index()
mr_emails_group = mr_emails_group[mr_emails_group['Quarter'].str.contains('Q')==True].reset_index(drop=True)
mr_emails_group['Order'] = mr_emails_group.Quarter.apply(lambda x: int(x.strip('Q')))
mr_emails_group = mr_emails_group.sort_values(['Year', 'Order'], ascending=[False, False]).reset_index(drop=True)
recent_qtr = str(mr_emails_group.iloc[0]['Quarter']) + ' ' + str(mr_emails_group.iloc[0]['Year'])
prev_qtr = str(mr_emails_group.iloc[1]['Quarter']) + ' ' + str(mr_emails_group.iloc[1]['Year'])
recent_mr_recipients = mr_emails_group.at[0,'Sent']
prev_mr_recipients = mr_emails_group.at[1,'Sent']
mr_recipients_change = int(((recent_mr_recipients - prev_mr_recipients)/prev_mr_recipients)*100)
fig13=plt.figure(figsize=(2.3,2.3))
ax13=fig13.add_subplot(111,aspect='equal')
ax13.add_artist(Wedge((.5,.5), 0.5, outer_theta1(mr_recipients_change), outer_theta2(mr_recipients_change), color=patch_color(mr_recipients_change), alpha=outer_alpha(mr_recipients_change)))
ax13.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(mr_recipients_change)))
ax13.add_artist(Wedge((.5,.5), 0.425, inner_theta1(mr_recipients_change), inner_theta2(mr_recipients_change), color=patch_color(mr_recipients_change), alpha=inner_alpha(mr_recipients_change)))
ax13.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax13.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig13.savefig('market-reports.eps',transparent=True)
recent_mr_recipients = '{:,}'.format(recent_mr_recipients)
prev_mr_recipients = '{:,}'.format(prev_mr_recipients)

# Advertising
current_loopnet_views = int(inputs_df.iloc[16]['CSV File Name or Numbers'])
prev_loopnet_views = int(inputs_df.iloc[17]['CSV File Name or Numbers'])
loopnet_views_change = int(((current_loopnet_views - prev_loopnet_views)/prev_loopnet_views)*100)
fig14=plt.figure(figsize=(2.3,2.3))
ax14=fig14.add_subplot(111,aspect='equal')
ax14.add_artist(Wedge((.5,.5), 0.5, outer_theta1(loopnet_views_change), outer_theta2(loopnet_views_change), color=patch_color(loopnet_views_change), alpha=outer_alpha(loopnet_views_change)))
ax14.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(loopnet_views_change)))
ax14.add_artist(Wedge((.5,.5), 0.425, inner_theta1(loopnet_views_change), inner_theta2(loopnet_views_change), color=patch_color(loopnet_views_change), alpha=inner_alpha(loopnet_views_change)))
ax14.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax14.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig14.savefig('ad-views.eps',transparent=True)
current_loopnet_views = '{:,}'.format(current_loopnet_views)
prev_loopnet_views = '{:,}'.format(prev_loopnet_views)

# Proposals
current_proposals_csv = inputs_df.iloc[18]['CSV File Name or Numbers']
while '.csv' not in current_proposals_csv:
    current_proposals_csv = input('Please provide the name of the CSV file, including the .csv, for this year\'s tracking report: ')
prev_proposals_csv = inputs_df.iloc[19]['CSV File Name or Numbers']
while '.csv' not in prev_proposals_csv:
    prev_proposals_csv = input('Please provide the name of the CSV file, including the .csv, for last year\'s tracking report: ')
current_proposals_df = pd.read_csv(current_proposals_csv)
current_proposals_df = current_proposals_df.fillna(value=1).reset_index(drop=True)
current_proposals_df['Submission_Month'] = current_proposals_df['Submission Date'].apply(lambda x: str(x).split('.')[0])
current_proposals_df['Submission_Month'] = current_proposals_df['Submission_Month'].apply(lambda x: int(x))
current_month = current_proposals_df['Submission_Month'].max()
prev_proposals_df = pd.read_csv(prev_proposals_csv)
prev_proposals_df = prev_proposals_df.fillna(value=1).reset_index(drop=True)
prev_proposals_df['Submission_Month'] = prev_proposals_df['Submission Date'].apply(lambda x: str(x).split('.')[0])
prev_proposals_df['Submission_Month'] = prev_proposals_df['Submission_Month'].apply(lambda x: int(x))
prev_proposals_ytd = prev_proposals_df[prev_proposals_df.Submission_Month <= current_month]
ytd_proposals = current_proposals_df['Submission_Month'].count()
prev_ytd_proposals = prev_proposals_ytd['Submission_Month'].count()
ytd_proposals_change = int(((ytd_proposals - prev_ytd_proposals)/prev_ytd_proposals)*100)
fig15=plt.figure(figsize=(2.3,2.3))
ax15=fig15.add_subplot(111,aspect='equal')
ax15.add_artist(Wedge((.5,.5), 0.5, outer_theta1(ytd_proposals_change), outer_theta2(ytd_proposals_change), color=patch_color(ytd_proposals_change), alpha=outer_alpha(ytd_proposals_change)))
ax15.add_artist(Circle((.5,.5), 0.435, color="#ffffff", alpha=outline_alpha(ytd_proposals_change)))
ax15.add_artist(Wedge((.5,.5), 0.425, inner_theta1(ytd_proposals_change), inner_theta2(ytd_proposals_change), color=patch_color(ytd_proposals_change), alpha=inner_alpha(ytd_proposals_change)))
ax15.add_artist(Circle((.5,.5),0.33,color="#ffffff",alpha=1))
ax15.add_artist(Circle((.5,.5),0.32,color="#00467f",alpha=1))
plt.axis('off')
plt.tight_layout()
fig15.savefig('proposals.eps',transparent=True)

# Get the file links for all the EPS files just created
adViewsEPS = cwd + '\\ad-views.eps'
devEmailsEPS = cwd + '\\development-emails.eps'
flexEmailsEPS = cwd + '\\flex-emails.eps'
indEmailsEPS = cwd + '\\industrial-emails.eps'
invEmailsEPS = cwd + '\\investment-emails.eps'
landEmailsEPS = cwd + '\\land-emails.eps'
marketReportsEPS = cwd + '\\market-reports.eps' 
medEmailsEPS = cwd + '\\medical-emails.eps'
offEmailsEPS = cwd + '\\office-emails.eps'
pressReleasesEPS = cwd + '\\press-releases.eps'
proposalsEPS = cwd + '\\proposals.eps'
retEmailsEPS = cwd + '\\retail-emails.eps'
socialImpEPS = cwd + '\\social-impressions.eps'
totalEmailsEPS = cwd + '\\total_emails_by_campaign.eps'
websiteVisitsEPS = cwd + '\\website-visits.eps'

# Get the file links for the social icons
twitterEPS = cwd + '\\twitter.eps'
facebookEPS = cwd + '\\facebook.eps'
instagramEPS = cwd + '\\instagram.eps'

# ALL THE INDESIGNNNNNNNNNN
# open InDesign and create a new document
app = win32com.client.Dispatch('InDesign.Application.CC.2018')
myDocument = app.Documents.Add()

# set up the page size and disable facing pages
myDocument.DocumentPreferences.PageHeight = "17i"
myDocument.DocumentPreferences.PageWidth = "11i"
myDocument.DocumentPreferences.PagesPerDocument = 1
myDocument.DocumentPreferences.FacingPages = False

# variables to force styling
myLinearGradient = 1635282023
myCenterAlign = 1667591796
myLeftAlign = 1818584692
myAllCaps = 1634493296
normalCaps = 1852797549
myGraphicType = 1735553140
strokeNone = myDocument.Swatches.Item("None")
myArrowHead = 1937203560
noArrowHead = 1852796517
myHeaderRow = 1162375799
myVAlignCenter = 1667591796
myVAlignBottom = 1651471469
myGraphicCellType = 1701728329
myFitProportionally = 1684885618

# set up colors
# Colliers Dark Blue
colliersDBlue = myDocument.Colors.Add()
colliersDBlue.Name = "Colliers Dark Blue"
colliersDBlue.ColorValue = [100, 57, 0, 38]
# Colliers Light Blue
colliersLBlue = myDocument.Colors.Add()
colliersLBlue.Name = "Colliers Light Blue"
colliersLBlue.ColorValue = [100, 10, 0, 10]
# 80% Gray
colliers80Gray = myDocument.Colors.Add()
colliers80Gray.Name = "Colliers 80 Gray"
colliers80Gray.ColorValue = [0, 0, 0, 80]
# Yellow
colliersYellow = myDocument.Colors.Add()
colliersYellow.Name = "Colliers Yellow"
colliersYellow.ColorValue = [0, 24, 94, 0]
# Red
colliersRed = myDocument.Colors.Add()
colliersRed.Name = "Colliers Red"
colliersRed.ColorValue = [0, 95, 100, 0]
# Dark Red
colliersDRed = myDocument.Colors.Add()
colliersDRed.Name = "Colliers Dark Red"
colliersDRed.ColorValue = [0, 95, 100, 29]
# White
colliersWhite = myDocument.Colors.Add()
colliersWhite.Name = "Colliers White"
colliersWhite.ColorValue = [0, 0, 0, 0]

# Create a page
myPage = myDocument.Pages.Item(1)

# set up paragraph styles
# fix [Basic Paragraph]
basicParagraph = myDocument.ParagraphStyles.Item("[Basic Paragraph]")
basicParagraph.FillColor = colliers80Gray
basicParagraph.PointSize = 9
basicParagraph.AppliedFont = app.Fonts.Item("Aaux Next")
basicParagraph.FontStyle = "Regular"
basicParagraph.Hyphenation = False
basicParagraph.Justification = myLeftAlign
basicParagraph.Capitalization = normalCaps

# Heading 1
heading1Style = myDocument.ParagraphStyles.Add()
heading1Style.Name = "Heading_1_Paragraph"
heading1Style.BasedOn = basicParagraph
heading1Style.PointSize = 21
heading1Style.FontStyle = "Ultra"
heading1Style.Capitalization = myAllCaps

# Heading 2
heading2Style = myDocument.ParagraphStyles.Add()
heading2Style.Name = "Heading_2_Paragraph"
heading2Style.BasedOn = basicParagraph
heading2Style.PointSize = 14
heading2Style.Capitalization = myAllCaps

# Subheading
subheadingStyle = myDocument.ParagraphStyles.Add()
subheadingStyle.Name = "Subheading"
subheadingStyle.BasedOn = basicParagraph
subheadingStyle.PointSize = 12
subheadingStyle.FontStyle = "Ultra"
subheadingStyle.Justification = myCenterAlign
subheadingStyle.Capitalization = myAllCaps
subheadingStyle.FillColor = colliersWhite

# Table Title
tableTitleStyle = myDocument.ParagraphStyles.Add()
tableTitleStyle.Name = "Table_Title"
tableTitleStyle.BasedOn = basicParagraph
tableTitleStyle.FontStyle = "SemiBold"
tableTitleStyle.Justification = myLeftAlign

# Table Headings
tableHeadingStyle = myDocument.ParagraphStyles.Add()
tableHeadingStyle.Name = "Table_Heading"
tableHeadingStyle.BasedOn = basicParagraph
tableHeadingStyle.FontStyle = "SemiBold"
tableHeadingStyle.Justification = myCenterAlign

# Table Data
tableDataStyle = myDocument.ParagraphStyles.Add()
tableDataStyle.Name = "Table_Data"
tableDataStyle.BasedOn = basicParagraph
tableDataStyle.Justification = myCenterAlign
tableDataStyle.Capitalization = normalCaps

# Table Headings Cell Style
tableHeadingCells = myDocument.CellStyles.Add()
tableHeadingCells.Name = 'Table Headings Cells'
tableHeadingCells.appliedParagraphStyle = tableHeadingStyle
tableHeadingCells.BottomEdgeStrokeColor = colliers80Gray
tableHeadingCells.BottomEdgeStrokeWeight = 1
tableHeadingCells.LeftEdgeStrokeColor = strokeNone
tableHeadingCells.LeftEdgeStrokeWeight = 0
tableHeadingCells.RightEdgeStrokeColor = strokeNone
tableHeadingCells.RightEdgeStrokeWeight = 0
tableHeadingCells.TopEdgeStrokeColor = strokeNone
tableHeadingCells.TopEdgeStrokeWeight = 0

# Basic Table Cell Style
basicCells = myDocument.CellStyles.Add()
basicCells.Name = 'Basic Headings Cells'
basicCells.appliedParagraphStyle = tableDataStyle
basicCells.BottomEdgeStrokeColor = colliers80Gray
basicCells.BottomEdgeStrokeWeight = 1
basicCells.LeftEdgeStrokeColor = strokeNone
basicCells.LeftEdgeStrokeWeight = 0
basicCells.RightEdgeStrokeColor = strokeNone
basicCells.RightEdgeStrokeWeight = 0
basicCells.TopEdgeStrokeColor = strokeNone
basicCells.TopEdgeStrokeWeight = 0
basicCells.OverprintFill = True
basicCells.ClipContentToTextCell = True
basicCells.ClipContentToGraphicCell = True

# Table Style
myTableStyle = myDocument.TableStyles.Add()
myTableStyle.Name = 'Dashboard Table'
myTableStyle.BottomBorderStrokeColor = strokeNone
myTableStyle.BottomBorderStrokeWeight = 0
myTableStyle.LeftBorderStrokeColor = strokeNone
myTableStyle.LeftBorderStrokeWeight = 0
myTableStyle.RightBorderStrokeColor = strokeNone
myTableStyle.RightBorderStrokeWeight = 0
myTableStyle.TopBorderStrokeColor = strokeNone
myTableStyle.TopBorderStrokeWeight = 0
myTableStyle.BodyRegionCellStyle = basicCells
myTableStyle.HeaderRegionCellStyle = tableHeadingCells
myTableStyle.GraphicBottomInset = '0.05i'
myTableStyle.GraphicTopInset = '0.05i'

# Social Object Style
socialIconObjectStyle = myDocument.ObjectStyles.Add()
socialIconObjectStyle.Name = 'Social Object'
socialIconObjectStyle.EnableFrameFittingOptions = True
socialIconObjectStyle.FrameFittingOptions.FittingOnEmptyFrame = 1668247152
socialIconObjectStyle.StrokeColor = strokeNone

# Circle Heading
circleHeadingStyle = myDocument.ParagraphStyles.Add()
circleHeadingStyle.Name = "Circle_Heading"
circleHeadingStyle.BasedOn = basicParagraph
circleHeadingStyle.FontStyle = "Ultra"
circleHeadingStyle.Justification = myCenterAlign
circleHeadingStyle.Capitalization = myAllCaps

# Circle Numbers
circleNumbersStyle = myDocument.ParagraphStyles.Add()
circleNumbersStyle.Name = "Circle_Numbers"
circleNumbersStyle.BasedOn = basicParagraph
circleNumbersStyle.PointSize = 16
circleNumbersStyle.FontStyle = "Ultra"
circleNumbersStyle.Justification = myCenterAlign
circleNumbersStyle.FillColor = colliersWhite

# Large Circle Numbers
largeCircleNumbersStyle = myDocument.ParagraphStyles.Add()
largeCircleNumbersStyle.Name = "Large_Circle_Numbers"
largeCircleNumbersStyle.BasedOn = circleNumbersStyle
largeCircleNumbersStyle.PointSize = 21

# Circle Text
circleTextStyle = myDocument.ParagraphStyles.Add()
circleTextStyle.Name = "Circle_Text"
circleTextStyle.BasedOn = basicParagraph
circleTextStyle.Justification = myCenterAlign
circleTextStyle.Capitalization = myAllCaps
circleTextStyle.FillColor = colliersWhite

# Percent Changed
percentChangedStyle = myDocument.ParagraphStyles.Add()
percentChangedStyle.Name = "Percent_Changed"
percentChangedStyle.BasedOn = basicParagraph
percentChangedStyle.PointSize = 16
percentChangedStyle.FontStyle = "Ultra"
percentChangedStyle.Justification = myCenterAlign
percentChangedStyle.FillColor = colliersDBlue

# Large Percent Change
largePercentChangeStyle = myDocument.ParagraphStyles.Add()
largePercentChangeStyle.Name = "Large_Percent_Change"
largePercentChangeStyle.BasedOn = percentChangedStyle
largePercentChangeStyle.PointSize = 21

# Data Text
dataTextStyle = myDocument.ParagraphStyles.Add()
dataTextStyle.Name = "Data_Text"
dataTextStyle.BasedOn = basicParagraph
dataTextStyle.Justification = myCenterAlign
dataTextStyle.Capitalization = myAllCaps

# Social Data
socialDataStyle = myDocument.ParagraphStyles.Add()
socialDataStyle.Name = "Social_Data"
socialDataStyle.BasedOn = basicParagraph
socialDataStyle.Capitalization = myAllCaps

# Negative Data Text
negativeDataTextStyle = myDocument.CharacterStyles.Add()
negativeDataTextStyle.Name = "Negative_Data_Text"
negativeDataTextStyle.FontStyle = "Black"
negativeDataTextStyle.FillColor = colliersDRed

# Nested GREP Style for negative text
negativeGREP = myDocument.ParagraphStyles.Item("Data_Text").nestedGrepStyles.Add()
negativeGREP.AppliedCharacterStyle = myDocument.CharacterStyles.Item("Negative_Data_Text")
negativeGREP.GrepExpression = "-\\d+\\.\\d+% Change"

# Compared Number
comparedNumberStyle = myDocument.ParagraphStyles.Add()
comparedNumberStyle.Name = "Compared_Number"
comparedNumberStyle.BasedOn = dataTextStyle
comparedNumberStyle.FontStyle = "Ultra"

# Plain Data
plainDataStyle = myDocument.ParagraphStyles.Add()
plainDataStyle.Name = "Plaim_Data"
plainDataStyle.BasedOn = basicParagraph
plainDataStyle.PointSize = 21
plainDataStyle.FontStyle = "Ultra"
plainDataStyle.Justification = myCenterAlign

# Plain Data Subheading
plainDataSubheadingStyle = myDocument.ParagraphStyles.Add()
plainDataSubheadingStyle.Name = "Email_Subheading"
plainDataSubheadingStyle.BasedOn = basicParagraph
plainDataSubheadingStyle.PointSize = 12
plainDataSubheadingStyle.Justification = myCenterAlign
plainDataSubheadingStyle.Capitalization = myAllCaps

# Arrow Direction Test
def left_arrow_test(perc):
    if perc < 0:
        return noArrowHead
    else:
        return myArrowHead
def right_arrow_test(perc):
    if perc < 0:
        return myArrowHead
    else:
        return noArrowHead

myDshbdTitle = myPage.TextFrames.Add()
myDshbdTitle.GeometricBounds = ["0.342i", "0.25i", "0.6515i", "3.7095i"]
myDshbdTitle.ParentStory.Contents = "Marketing Dashboard"
titleText = myDshbdTitle.ParentStory.Characters.Item(1)
titleText.appliedParagraphStyle = heading1Style

myDshbdMonth = myPage.TextFrames.Add()
myDshbdMonth.GeometricBounds = ["0.6515i", "0.25i", "0.9462i", "3.7095i"]
myDshbdMonth.ParentStory.Contents = "South Carolina | " + recent_month_name + " 2018"
monthText = myDshbdMonth.ParentStory.Characters.Item(1)
monthText.appliedParagraphStyle = heading2Style

# Social Media Subheading
socialRect = myPage.Rectangles.Add()
socialRect.GeometricBounds = ["1.084i","0.25i","1.4211i","5.5i"]
socialRect.FillColor = colliersDBlue
socialRect.StrokeColor = strokeNone
socialSubH = myPage.TextFrames.Add()
socialSubH.GeometricBounds = ["1.1944i","1.3477i","1.335i","4.4023i"]
socialSubH.ParentStory.Contents = "Social Media"
socialSubHText = socialSubH.ParentStory.Characters.Item(1)
socialSubHText.appliedParagraphStyle = subheadingStyle

# Website Visits Subheading
webRect = myPage.Rectangles.Add()
webRect.GeometricBounds = ["0.25i","5.5i","0.5871i","10.75i"]
webRect.FillColor = colliersDBlue
webRect.StrokeColor = strokeNone
webSubH = myPage.TextFrames.Add()
webSubH.GeometricBounds = ["0.3604i","6.5977i","0.5i","9.6523i"]
webSubH.ParentStory.Contents = "Colliers.com Website Visits"
webSubHText = webSubH.ParentStory.Characters.Item(1)
webSubHText.appliedParagraphStyle = subheadingStyle

# Marketing Initiatives Subheading
mktgRect = myPage.Rectangles.Add()
mktgRect.GeometricBounds = ["5.9659i","0.25i","6.302i","10.75i"]
mktgRect.FillColor = colliersDBlue
mktgRect.StrokeColor = strokeNone
mktgSubH = myPage.TextFrames.Add()
mktgSubH.GeometricBounds = ["6.0753i","3.9727i","6.208i","7.0272i"]
mktgSubH.ParentStory.Contents = "Marketing Initiatives"
mktgSubHText = mktgSubH.ParentStory.Characters.Item(1)
mktgSubHText.appliedParagraphStyle = subheadingStyle

# Company-Wide Emails Subheading
emailsRect = myPage.Rectangles.Add()
emailsRect.GeometricBounds = ["8.7025i","0.25i","9.0396i","5.5i"]
emailsRect.FillColor = colliersDBlue
emailsRect.StrokeColor = strokeNone
emailsSubH = myPage.TextFrames.Add()
emailsSubH.GeometricBounds = ["8.8129i","1.3477i","8.96i","4.4023i"]
emailsSubH.ParentStory.Contents = "Company-wide Emails"
emailsSubHText = emailsSubH.ParentStory.Characters.Item(1)
emailsSubHText.appliedParagraphStyle = subheadingStyle

# Social Media impressions
socialGraphic = myPage.Place(socialImpEPS)
socialGraphic = socialGraphic.Item(1)
socialFrame = socialGraphic.Parent
socialFrame.GeometricBounds = ["1.2829i","-0.0883i","3.5746i","2.2033i"]
socialGraphic.GeometricBounds = ["1.2829i","-0.0883i","3.5746i","2.2033i"]

socialCirNum = myPage.TextFrames.Add()
socialCirNum.GeometricBounds = ["2.1148i","0.7405i","2.3602i","1.5432i"]
socialCirNum.ParentStory.Contents = recent_impressions_total
socialCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

socialCirText = myPage.TextFrames.Add()
socialCirText.GeometricBounds = ["2.3591i","0.7405i","2.6182i","1.5432i"]
socialCirText.ParentStory.Contents = recent_mo + "\nImpressions"
socialCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

socialArrow = myPage.GraphicLines.Add()
socialArrow.GeometricBounds = ["1.7i","2.4145i","1.9578i","2.4145i"]
socialArrow.StrokeColor = colliersDBlue
socialArrow.StrokeWeight = 4
socialArrow.LeftLineEnd = left_arrow_test(impressions_perc_change)
socialArrow.RightLineEnd = right_arrow_test(impressions_perc_change)

socialPercChanged = myPage.TextFrames.Add()
socialPercChanged.GeometricBounds = ["2.0542i","2.0995i","2.2976i","2.7295i"]
socialPercChanged.ParentStory.Contents = str(impressions_perc_change) + "%"
socialPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

socialComText = myPage.TextFrames.Add()
socialComText.GeometricBounds = ["2.2976i","2.0995i","2.7715i","2.7295i"]
socialComText.ParentStory.Contents = "Compared to " + prev_mo + " Total"
socialComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

socialComNum = myPage.TextFrames.Add()
socialComNum.GeometricBounds = ["2.7715i","2.0995i","2.9147i","2.7295i"]
socialComNum.ParentStory.Contents = prev_impressions_total
socialComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

twitterIcon = myPage.Place(twitterEPS)
twitterIcon = twitterIcon.Item(1)
twitterIconFrame = twitterIcon.Parent
twitterIconFrame.GeometricBounds = ["1.841i","3.3166i","2.1755i","3.6511i"]
twitterIcon.GeometricBounds = ["1.841i","3.3166i","2.1755i","3.6511i"]
twitterData = myPage.TextFrames.Add()
twitterData.GeometricBounds = ["1.8703i","3.8065i","2.1339i","5.4769i"]
twitterData.ParentStory.Contents = "Followers | "+ count_twitter_followers + "\nPosts | " + count_twitter_posts
twitterData.ParentStory.Characters.Item(1).appliedParagraphStyle = socialDataStyle

facebookIcon = myPage.Place(facebookEPS)
facebookIcon = facebookIcon.Item(1)
facebookIconFrame = facebookIcon.Parent
facebookIconFrame.GeometricBounds = ["2.3191i","3.3166i","2.6536i","3.6511i"]
facebookIcon.GeometricBounds = ["2.3191i","3.3166i","2.6536i","3.6511i"]
facebookData = myPage.TextFrames.Add()
facebookData.GeometricBounds = ["2.3546i","3.8065i","2.6182i","5.4769i"]
facebookData.ParentStory.Contents = "Followers | "+ count_fb_followers + "\nPosts | " + count_fb_posts
facebookData.ParentStory.Characters.Item(1).appliedParagraphStyle = socialDataStyle

instagramIcon = myPage.Place(instagramEPS)
instagramIcon = instagramIcon.Item(1)
instagramIconFrame = instagramIcon.Parent
instagramIconFrame.GeometricBounds = ["2.7715i","3.3166i","3.106i","3.6511i"]
instagramIcon.GeometricBounds = ["2.7715i","3.3166i","3.106i","3.6511i"]
instagramData = myPage.TextFrames.Add()
instagramData.GeometricBounds = ["2.8021i","3.8065i","3.0657i","5.4769i"]
instagramData.ParentStory.Contents = "Followers | "+ count_insta_followers + "\nPosts | " + count_insta_posts
instagramData.ParentStory.Characters.Item(1).appliedParagraphStyle = socialDataStyle

socialTableText = myPage.TextFrames.Add()
socialTableText.GeometricBounds = ["3.3735i","0.26i","5.7768i","5.6i"]
socialTableText.ParentStory.Contents = "Top Performing Social Posts"
socialTableText.ParentStory.Characters.Item(1).appliedParagraphStyle = tableTitleStyle
socialTable = socialTableText.ParentStory.InsertionPoints.Item(-1).Tables.Add()
socialcolumncount = 5
socialrowcount = 6
socialTable.ColumnCount = socialcolumncount
socialTable.HeaderRowCount = 1
socialTable.BodyRowCount = socialrowcount - 1
socialTable.Height = '2.0i'
socialTable.appliedTableStyle = myTableStyle
socialTable.Columns.Item(1).Width = '2.7319i'
socialTable.Columns.Item(2).Width = '0.6592i'
socialTable.Columns.Item(3).Width = '0.7419i'
socialTable.Columns.Item(4).Width = '0.455i'
socialTable.Columns.Item(5).Width = '0.733i'
socialTable.Columns.Item(4).FillColor = colliersYellow
socialTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
socialTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
socialTable.Rows.Item(1).VerticalJustification = myVAlignBottom
for i in range(2,7):
    socialTable.Rows.Item(i).VerticalJustification = myVAlignCenter
for i in range(1,7):
    socialTable.Rows.Item(i).Height = '0.375i'
    socialTable.Rows.Item(i).AutoGrow = False
    socialTable.Rows.Item(i).ClipContentToTextCell = True
    socialTable.Rows.Item(i).OverprintFill = True
socialTable.Contents = top_social_posts_list
socialrangetotal = socialcolumncount * socialrowcount
for i in range(1,socialrangetotal,socialcolumncount):
    socialTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign
for i in range(7,socialrangetotal,socialcolumncount):
    if socialTable.Cells.Item(i).Contents == 'Facebook':
        socialTable.Cells.Item(i).Contents = ''
        socialTable.Cells.Item(i).ConvertCellType(myGraphicCellType)
        socialTable.Cells.Item(i).ClipContentToGraphicCell = True
        socialIconFrame = socialTable.Cells.Item(i).Rectangles.Add()
        socialIconFrame.appliedObjectStyle = socialIconObjectStyle
        socialIconImage = socialIconFrame.Place(facebookEPS)
    elif socialTable.Cells.Item(i).Contents == 'Twitter':
        socialTable.Cells.Item(i).Contents = ''
        socialTable.Cells.Item(i).ConvertCellType(myGraphicCellType)
        socialTable.Cells.Item(i).ClipContentToGraphicCell = True
        socialIconFrame = socialTable.Cells.Item(i).Rectangles.Add()
        socialIconFrame.appliedObjectStyle = socialIconObjectStyle
        socialIconImage = socialIconFrame.Place(twitterEPS)
    else:
        socialTable.Cells.Item(i).Contents = ''
        socialTable.Cells.Item(i).ConvertCellType(myGraphicCellType)
        socialTable.Cells.Item(i).ClipContentToGraphicCell = True
        socialIconFrame = socialTable.Cells.Item(i).Rectangles.Add()
        socialIconFrame.appliedObjectStyle = socialIconObjectStyle
        socialIconImage = socialIconFrame.Place(instagramEPS)

# Website visits
websiteGraphic = myPage.Place(websiteVisitsEPS)
websiteGraphic = websiteGraphic.Item(1)
websiteFrame = websiteGraphic.Parent
websiteFrame.GeometricBounds = ["0.5422i","5.3479i","3.6394i","8.4451i"]
websiteGraphic.GeometricBounds = ["0.5422i","5.3479i","3.6394i","8.4451i"]

websiteCirNum = myPage.TextFrames.Add()
websiteCirNum.GeometricBounds = ["1.7259i","6.3318i","2.0021i","7.489i"]
websiteCirNum.ParentStory.Contents = total_page_views
websiteCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = largeCircleNumbersStyle

websiteCirText = myPage.TextFrames.Add()
websiteCirText.GeometricBounds = ["2.0846i","6.3318i","2.4864i","7.489i"]
websiteCirText.ParentStory.Contents = recent_mo + "\nTotal Website Visits"
websiteCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

campaignTableText = myPage.TextFrames.Add()
campaignTableText.GeometricBounds = ["0.7762i","8.2787i","4.0896i","10.6213i"]
campaignTableText.ParentStory.Contents = "Visits from Campaigns"
campaignTableText.ParentStory.Characters.Item(1).appliedParagraphStyle = tableTitleStyle
campaignTable = campaignTableText.ParentStory.InsertionPoints.Item(-1).Tables.Add()
campaigncolumncount = 2
campaignrowcount = len(campaign_visits_group) + 1
campaignTable.ColumnCount = campaigncolumncount
campaignTable.HeaderRowCount = 1
campaignTable.BodyRowCount = campaignrowcount - 1
campaignTable.Height = '3i'
campaignTable.appliedTableStyle = myTableStyle
campaignTable.Columns.Item(1).Width = '1.5502i'
campaignTable.Columns.Item(2).Width = '0.75i'
campaignTable.Columns.Item(2).FillColor = colliersYellow
campaignTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
campaignTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
campaignTable.Rows.Item(1).VerticalJustification = myVAlignBottom
campaignTable.Contents = campaign_visits_list
campaignrangetotal = campaigncolumncount * campaignrowcount
for i in range(2,campaignrowcount):
    campaignTable.Rows.Item(i).VerticalJustification = myVAlignCenter
for i in range(1,campaignrangetotal,campaigncolumncount):
    campaignTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign

pagesTableText = myPage.TextFrames.Add()
pagesTableText.GeometricBounds = ["3.5466i","5.7598i","5.7768i","8.1096i"]
pagesTableText.ParentStory.Contents = "Which pages were they visiting?"
pagesTableText.ParentStory.Characters.Item(1).appliedParagraphStyle = tableTitleStyle
pagesTable = pagesTableText.ParentStory.InsertionPoints.Item(-1).Tables.Add()
pagescolumncount = 2
pagesrowcount = len(page_views_table) + 1
pagesTable.ColumnCount = pagescolumncount
pagesTable.HeaderRowCount = 1
pagesTable.BodyRowCount = pagesrowcount - 1
pagesTable.Height = '2i'
pagesTable.appliedTableStyle = myTableStyle
pagesTable.Columns.Item(1).Width = '1.5356i'
pagesTable.Columns.Item(2).Width = '0.8142i'
pagesTable.Columns.Item(2).FillColor = colliersYellow
pagesTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
pagesTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
pagesTable.Rows.Item(1).VerticalJustification = myVAlignBottom
for i in range(2,7):
    pagesTable.Rows.Item(i).VerticalJustification = myVAlignCenter
pagesTable.Contents = page_views_table_list
pagesrangetotal = pagescolumncount * pagesrowcount
for i in range(1,pagesrangetotal,pagescolumncount):
    pagesTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign

propSubText = myPage.TextFrames.Add()
propSubText.GeometricBounds = ["4.287i","8.5313i","4.6215i","10.3688i"]
propSubText.ParentStory.Contents = current_prop_subscribers
propSubText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
propSubHdg = myPage.TextFrames.Add()
propSubHdg.GeometricBounds = ["4.62i","8.5313i","4.9571i","10.3688i"]
propSubHdg.ParentStory.Contents = "Property Subscriber Leads"
propSubHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

resSubText = myPage.TextFrames.Add()
resSubText.GeometricBounds = ["5.1066i","8.5313i","5.4412i","10.3688i"]
resSubText.ParentStory.Contents = current_research_subscribers
resSubText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
resSubHdg = myPage.TextFrames.Add()
resSubHdg.GeometricBounds = ["5.4397i","8.5313i","5.7768i","10.3688i"]
resSubHdg.ParentStory.Contents = "Research Subscriber Leads"
resSubHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

# Press Releases
prGraphic = myPage.Place(pressReleasesEPS)
prGraphic = prGraphic.Item(1)
prFrame = prGraphic.Parent
prFrame.GeometricBounds = ["6.2222i","-0.1439i","8.5139i","2.1478i"]
prGraphic.GeometricBounds = ["6.2222i","-0.1439i","8.5139i","2.1478i"]

prCirNum = myPage.TextFrames.Add()
prCirNum.GeometricBounds = ["7.0571i","0.6915i","7.3025i","1.4943i"]
prCirNum.ParentStory.Contents = str(ytd_prs)
prCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

prCirText = myPage.TextFrames.Add()
prCirText.GeometricBounds = ["7.315i","0.6915i","7.5659i","1.4943i"]
prCirText.ParentStory.Contents = "YTD 2018 Released"
prCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

prCirHdg = myPage.TextFrames.Add()
prCirHdg.GeometricBounds = ["8.2337i","0.2079i","8.375i","1.9779i"]
prCirHdg.ParentStory.Contents = "Press Releases"
prCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

prArrow = myPage.GraphicLines.Add()
prArrow.GeometricBounds = ["6.6423i","2.3656i","6.9001i","2.3656i"]
prArrow.StrokeColor = colliersDBlue
prArrow.StrokeWeight = 4
prArrow.LeftLineEnd = left_arrow_test(prs_perc_change)
prArrow.RightLineEnd = right_arrow_test(prs_perc_change)

prPercChanged = myPage.TextFrames.Add()
prPercChanged.GeometricBounds = ["6.9965i","2.0505i","7.2399i","2.6806i"]
prPercChanged.ParentStory.Contents = str(prs_perc_change) + "%"
prPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

prComText = myPage.TextFrames.Add()
prComText.GeometricBounds = ["7.2399i","2.0505i","7.7138i","2.6806i"]
prComText.ParentStory.Contents = "Compared to YTD 2017"
prComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

prComNum = myPage.TextFrames.Add()
prComNum.GeometricBounds = ["7.7138i","2.0505i","7.857i","2.6806i"]
prComNum.ParentStory.Contents = str(prev_ytd_prs)
prComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Market Reports
mrGraphic = myPage.Place(marketReportsEPS)
mrGraphic = mrGraphic.Item(1)
mrFrame = mrGraphic.Parent
mrFrame.GeometricBounds = ["6.2222i","2.5255i","8.5139i","4.8171i"]
mrGraphic.GeometricBounds = ["6.2222i","2.5255i","8.5139i","4.8171i"]

mrCirNum = myPage.TextFrames.Add()
mrCirNum.GeometricBounds = ["6.9965i","3.3582i","7.2419i","4.161i"]
mrCirNum.ParentStory.Contents = recent_mr_recipients
mrCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

mrCirText = myPage.TextFrames.Add()
mrCirText.GeometricBounds = ["7.2408i","3.3582i","7.64i","4.161i"]
mrCirText.ParentStory.Contents = recent_qtr + " Total Email Recipients"
mrCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

mrCirHdg = myPage.TextFrames.Add()
mrCirHdg.GeometricBounds = ["8.2337i","2.8746i","8.375i","4.6446i"]
mrCirHdg.ParentStory.Contents = "Market Reports"
mrCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

mrArrow = myPage.GraphicLines.Add()
mrArrow.GeometricBounds = ["6.6423i","5.0322i","6.9001i","5.0322i"]
mrArrow.StrokeColor = colliersDBlue
mrArrow.StrokeWeight = 4
mrArrow.LeftLineEnd = left_arrow_test(mr_recipients_change)
mrArrow.RightLineEnd = right_arrow_test(mr_recipients_change)

mrPercChanged = myPage.TextFrames.Add()
mrPercChanged.GeometricBounds = ["6.9965i","4.7172i","7.2399i","5.3472i"]
mrPercChanged.ParentStory.Contents = str(mr_recipients_change) + "%"
mrPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

mrComText = myPage.TextFrames.Add()
mrComText.GeometricBounds = ["7.2399i","4.7172i","7.7138i","5.3472i"]
mrComText.ParentStory.Contents = "Compared to\n" + prev_qtr
mrComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

mrComNum = myPage.TextFrames.Add()
mrComNum.GeometricBounds = ["7.7138i","4.7172i","7.857i","5.3472i"]
mrComNum.ParentStory.Contents = prev_mr_recipients
mrComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Ad Views
adViewsGraphic = myPage.Place(adViewsEPS)
adViewsGraphic = adViewsGraphic.Item(1)
adViewsFrame = adViewsGraphic.Parent
adViewsFrame.GeometricBounds = ["6.2222i","5.1894i","8.5139i","7.4811i"]
adViewsGraphic.GeometricBounds = ["6.2222i","5.1894i","8.5139i","7.4811i"]

adViewsCirNum = myPage.TextFrames.Add()
adViewsCirNum.GeometricBounds = ["7.0571i","6.0249i","7.3025i","6.8276i"]
adViewsCirNum.ParentStory.Contents = current_loopnet_views
adViewsCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

adViewsCirText = myPage.TextFrames.Add()
adViewsCirText.GeometricBounds = ["7.315i","6.0249i","7.5659i","6.8276i"]
adViewsCirText.ParentStory.Contents = "LoopNet Ad Views"
adViewsCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

adViewsCirHdg = myPage.TextFrames.Add()
adViewsCirHdg.GeometricBounds = ["8.2337i","5.5412i","8.375i","7.3112i"]
adViewsCirHdg.ParentStory.Contents = "Advertising"
adViewsCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

adViewsArrow = myPage.GraphicLines.Add()
adViewsArrow.GeometricBounds = ["6.6423i","7.6989i","6.9001i","7.6989i"]
adViewsArrow.StrokeColor = colliersDBlue
adViewsArrow.StrokeWeight = 4
adViewsArrow.LeftLineEnd = left_arrow_test(loopnet_views_change)
adViewsArrow.RightLineEnd = right_arrow_test(loopnet_views_change)

adViewsPercChanged = myPage.TextFrames.Add()
adViewsPercChanged.GeometricBounds = ["6.9965i","7.3839i","7.2399i","8.0139i"]
adViewsPercChanged.ParentStory.Contents = str(loopnet_views_change) + "%"
adViewsPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

adViewsComText = myPage.TextFrames.Add()
adViewsComText.GeometricBounds = ["7.2399i","7.3839i","7.7138i","8.0139i"]
adViewsComText.ParentStory.Contents = "Compared to " + prev_mo + " Ad Views"
adViewsComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

adViewsComNum = myPage.TextFrames.Add()
adViewsComNum.GeometricBounds = ["7.7138i","7.3839i","7.857i","8.0139i"]
adViewsComNum.ParentStory.Contents = prev_loopnet_views
adViewsComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Proposals
proposalsGraphic = myPage.Place(proposalsEPS)
proposalsGraphic = proposalsGraphic.Item(1)
proposalsFrame = proposalsGraphic.Parent
proposalsFrame.GeometricBounds = ["6.2222i","7.8581i","8.5139i","10.1498i"]
proposalsGraphic.GeometricBounds = ["6.2222i","7.8581i","8.5139i","10.1498i"]

proposalsCirNum = myPage.TextFrames.Add()
proposalsCirNum.GeometricBounds = ["7.0571i","8.6915i","7.3025i","9.4943i"]
proposalsCirNum.ParentStory.Contents = str(ytd_proposals)
proposalsCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

proposalsCirText = myPage.TextFrames.Add()
proposalsCirText.GeometricBounds = ["7.315i","8.6915i","7.5659i","9.4943i"]
proposalsCirText.ParentStory.Contents = "YTD Submitted"
proposalsCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

proposalsCirHdg = myPage.TextFrames.Add()
proposalsCirHdg.GeometricBounds = ["8.2337i","8.2079i","8.375i","9.9779i"]
proposalsCirHdg.ParentStory.Contents = "Proposals"
proposalsCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

proposalsArrow = myPage.GraphicLines.Add()
proposalsArrow.GeometricBounds = ["6.6423i","10.3656i","6.9001i","10.3656i"]
proposalsArrow.StrokeColor = colliersDBlue
proposalsArrow.StrokeWeight = 4
proposalsArrow.LeftLineEnd = left_arrow_test(ytd_proposals_change)
proposalsArrow.RightLineEnd = right_arrow_test(ytd_proposals_change)

proposalsPercChanged = myPage.TextFrames.Add()
proposalsPercChanged.GeometricBounds = ["6.9965i","10.0505i","7.2399i","10.6806i"]
proposalsPercChanged.ParentStory.Contents = str(ytd_proposals_change) + "%"
proposalsPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

proposalsComText = myPage.TextFrames.Add()
proposalsComText.GeometricBounds = ["7.2399i","10.0505i","7.7138i","10.6806i"]
proposalsComText.ParentStory.Contents = "Compared to YTD 2017"
proposalsComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

proposalsComNum = myPage.TextFrames.Add()
proposalsComNum.GeometricBounds = ["7.7138i","10.0505i","7.857i","10.6806i"]
proposalsComNum.ParentStory.Contents = str(prev_ytd_proposals)
proposalsComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Company Emails
opensText = myPage.TextFrames.Add()
opensText.GeometricBounds = ["9.1702i","0.2292i","9.5047i","1.3234i"]
opensText.ParentStory.Contents = total_opens
opensText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
opensHdg = myPage.TextFrames.Add()
opensHdg.GeometricBounds = ["9.5033i","0.2292i","9.66i","1.3234i"]
opensHdg.ParentStory.Contents = "Opens"
opensHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

openRateText = myPage.TextFrames.Add()
openRateText.GeometricBounds = ["9.1702i","1.6133i","9.5047i","2.7075i"]
openRateText.ParentStory.Contents = str(total_open_rate) + "%"
openRateText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
openRateHdg = myPage.TextFrames.Add()
openRateHdg.GeometricBounds = ["9.5033i","1.6133i","9.66i","2.7075i"]
openRateHdg.ParentStory.Contents = "Open Rate"
openRateHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

clicksText = myPage.TextFrames.Add()
clicksText.GeometricBounds = ["9.1702i","2.9974i","9.5047i","4.0916i"]
clicksText.ParentStory.Contents = total_clicks
clicksText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
clicksHdg = myPage.TextFrames.Add()
clicksHdg.GeometricBounds = ["9.5033i","2.9974i","9.66i","4.0916i"]
clicksHdg.ParentStory.Contents = "Clicks"
clicksHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

clickRateText = myPage.TextFrames.Add()
clickRateText.GeometricBounds = ["9.1702i","4.3814i","9.5047i","5.4757i"]
clickRateText.ParentStory.Contents = str(total_click_rate) + "%"
clickRateText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
clickRateHdg = myPage.TextFrames.Add()
clickRateHdg.GeometricBounds = ["9.5033i","4.3814i","9.66i","5.4757i"]
clickRateHdg.ParentStory.Contents = "Click Rate"
clickRateHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

emailsTableText = myPage.TextFrames.Add()
emailsTableText.GeometricBounds = ["9.894i","0.25i","12.1819i","6.107i"]
emailsTableText.ParentStory.Contents = "Top Performing Property/Campaign Emails"
emailsTableText.ParentStory.Characters.Item(1).appliedParagraphStyle = tableTitleStyle
emailsTable = emailsTableText.ParentStory.InsertionPoints.Item(-1).Tables.Add()
emailscolumncount = 7
emailsrowcount = 6
emailsTable.ColumnCount = emailscolumncount
emailsTable.HeaderRowCount = 1
emailsTable.BodyRowCount = emailsrowcount - 1
emailsTable.Height = '1.9i'
emailsTable.appliedTableStyle = myTableStyle
emailsTable.Columns.Item(1).Width = '1.9312i'
emailsTable.Columns.Item(2).Width = '0.7935i'
emailsTable.Columns.Item(3).Width = '0.535i'
emailsTable.Columns.Item(4).Width = '0.5712i'
emailsTable.Columns.Item(5).Width = '0.5288i'
emailsTable.Columns.Item(6).Width = '0.9083i'
emailsTable.Columns.Item(7).Width = '0.589i'
emailsTable.Columns.Item(5).FillColor = colliersYellow
emailsTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
emailsTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
emailsTable.Rows.Item(1).VerticalJustification = myVAlignBottom
for i in range(2,7):
    emailsTable.Rows.Item(i).VerticalJustification = myVAlignCenter
emailsTable.Contents = top_emails_list
emailsrangetotal = emailscolumncount * emailsrowcount
for i in range(1,emailsrangetotal,emailscolumncount):
    emailsTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign

companyEmailsGraphic = myPage.Place(totalEmailsEPS)
companyEmailsGraphic = companyEmailsGraphic.Item(1)
companyEmailsFrame = companyEmailsGraphic.Parent
companyEmailsFrame.GeometricBounds = ["8.45i","6.25i","12.45i","10.25i"]
companyEmailsGraphic.GeometricBounds = ["8.45i","6.25i","12.45i","10.25i"]

companyEmailsCirNum = myPage.TextFrames.Add()
companyEmailsCirNum.GeometricBounds = ["10.0912i","7.6768i","10.4131i","8.834i"]
companyEmailsCirNum.ParentStory.Contents = total_sent_emails
companyEmailsCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = largeCircleNumbersStyle

companyEmailsCirText = myPage.TextFrames.Add()
companyEmailsCirText.GeometricBounds = ["10.4499i","7.8581i","10.8795i","8.6527i"]
companyEmailsCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
companyEmailsCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

companyEmailsArrow = myPage.GraphicLines.Add()
companyEmailsArrow.GeometricBounds = ["9.5926i","10.3334i","9.9896i","10.3334i"]
companyEmailsArrow.StrokeColor = colliersDBlue
companyEmailsArrow.StrokeWeight = 7
companyEmailsArrow.LeftLineEnd = left_arrow_test(sent_perc)
companyEmailsArrow.RightLineEnd = right_arrow_test(sent_perc)

companyEmailsPercChanged = myPage.TextFrames.Add()
companyEmailsPercChanged.GeometricBounds = ["10.0912i","9.9169i","10.4131i","10.75i"]
companyEmailsPercChanged.ParentStory.Contents = str(sent_perc) + "%"
companyEmailsPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = largePercentChangeStyle

companyEmailsComText = myPage.TextFrames.Add()
companyEmailsComText.GeometricBounds = ["10.4131i","9.9169i","10.8795i","10.75i"]
companyEmailsComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
companyEmailsComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

companyEmailsComNum = myPage.TextFrames.Add()
companyEmailsComNum.GeometricBounds = ["10.8795i","9.9169i","11.0689i","10.75i"]
companyEmailsComNum.ParentStory.Contents = total_prev_sent_emails
companyEmailsComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Development Emails
devEmailsGraphic = myPage.Place(devEmailsEPS)
devEmailsGraphic = devEmailsGraphic.Item(1)
devEmailsFrame = devEmailsGraphic.Parent
devEmailsFrame.GeometricBounds = ["12.0955i","-0.1439i","14.3871i","2.1478i"]
devEmailsGraphic.GeometricBounds = ["12.0955i","-0.1439i","14.3871i","2.1478i"]

devCirNum = myPage.TextFrames.Add()
devCirNum.GeometricBounds = ["12.8552i","0.6915i","13.1007i","1.4943i"]
devCirNum.ParentStory.Contents = dev_emails
devCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

devCirText = myPage.TextFrames.Add()
devCirText.GeometricBounds = ["13.0996i","0.6915i","13.5i","1.4943i"]
devCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
devCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

devCirHdg = myPage.TextFrames.Add()
devCirHdg.GeometricBounds = ["14.0924i","0.1554i","14.2337i","2.0304i"]
devCirHdg.ParentStory.Contents = "Development"
devCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

devDataText = myPage.TextFrames.Add()
devDataText.GeometricBounds = ["14.2337i","0.1554i","14.4934i","2.0304i"]
devDataText.ParentStory.Contents = str(dev_open_rate) + "% Opens | " + str(dev_open_change) + "% Change\n" + str(dev_click_rate) + "% Clicks | " + str(dev_click_change) + "% Change"
devDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

devArrow = myPage.GraphicLines.Add()
devArrow.GeometricBounds = ["12.5011i","2.3656i","12.7588i","2.3656i"]
devArrow.StrokeColor = colliersDBlue
devArrow.StrokeWeight = 4
devArrow.LeftLineEnd = left_arrow_test(dev_perc)
devArrow.RightLineEnd = right_arrow_test(dev_perc)

devPercChanged = myPage.TextFrames.Add()
devPercChanged.GeometricBounds = ["12.8552i","2.0218i","13.0986i","2.7093i"]
devPercChanged.ParentStory.Contents = str(dev_perc) + "%"
devPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

devComText = myPage.TextFrames.Add()
devComText.GeometricBounds = ["13.0986i","2.0218i","13.6767i","2.7093i"]
devComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
devComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

devComNum = myPage.TextFrames.Add()
devComNum.GeometricBounds = ["13.6767i","2.0218i","13.8199i","2.7093i"]
devComNum.ParentStory.Contents = dev_prev_emails
devComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Flex Emails
flexEmailsGraphic = myPage.Place(flexEmailsEPS)
flexEmailsGraphic = flexEmailsGraphic.Item(1)
flexEmailsFrame = flexEmailsGraphic.Parent
flexEmailsFrame.GeometricBounds = ["12.0955i","2.5255i","14.3871i","4.8171i"]
flexEmailsGraphic.GeometricBounds = ["12.0955i","2.5255i","14.3871i","4.8171i"]

flexCirNum = myPage.TextFrames.Add()
flexCirNum.GeometricBounds = ["12.8552i","3.3582i","13.1007i","4.161i"]
flexCirNum.ParentStory.Contents = flex_emails
flexCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

flexCirText = myPage.TextFrames.Add()
flexCirText.GeometricBounds = ["13.0996i","3.3582i","13.5i","4.161i"]
flexCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
flexCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

flexCirHdg = myPage.TextFrames.Add()
flexCirHdg.GeometricBounds = ["14.0924i","2.8221i","14.2337i","4.6971i"]
flexCirHdg.ParentStory.Contents = "Flex"
flexCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

flexDataText = myPage.TextFrames.Add()
flexDataText.GeometricBounds = ["14.2337i","2.8221i","14.4934i","4.6971i"]
flexDataText.ParentStory.Contents = str(flex_open_rate) + "% Opens | " + str(flex_open_change) + "% Change\n" + str(flex_click_rate) + "% Clicks | " + str(flex_click_change) + "% Change"
flexDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

flexArrow = myPage.GraphicLines.Add()
flexArrow.GeometricBounds = ["12.5011i","5.0322i","12.7588i","5.0322i"]
flexArrow.StrokeColor = colliersDBlue
flexArrow.StrokeWeight = 4
flexArrow.LeftLineEnd = left_arrow_test(flex_perc)
flexArrow.RightLineEnd = right_arrow_test(flex_perc)

flexPercChanged = myPage.TextFrames.Add()
flexPercChanged.GeometricBounds = ["12.8552i","4.6885i","13.0986i","5.376i"]
flexPercChanged.ParentStory.Contents = str(flex_perc) + "%"
flexPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

flexComText = myPage.TextFrames.Add()
flexComText.GeometricBounds = ["13.0986i","4.6885i","13.6767i","5.376i"]
flexComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
flexComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

flexComNum = myPage.TextFrames.Add()
flexComNum.GeometricBounds = ["13.6767i","4.6885i","13.8199i","5.376i"]
flexComNum.ParentStory.Contents = flex_prev_emails
flexComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Healthcare Emails
medEmailsGraphic = myPage.Place(medEmailsEPS)
medEmailsGraphic = medEmailsGraphic.Item(1)
medEmailsFrame = medEmailsGraphic.Parent
medEmailsFrame.GeometricBounds = ["12.0955i","5.1894i","14.3871i","7.4811i"]
medEmailsGraphic.GeometricBounds = ["12.0955i","5.1894i","14.3871i","7.4811i"]

medCirNum = myPage.TextFrames.Add()
medCirNum.GeometricBounds = ["12.8552i","6.0249i","13.1007i","6.8276i"]
medCirNum.ParentStory.Contents = med_emails
medCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

medCirText = myPage.TextFrames.Add()
medCirText.GeometricBounds = ["13.0996i","6.0249i","13.5i","6.8276i"]
medCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
medCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

medCirHdg = myPage.TextFrames.Add()
medCirHdg.GeometricBounds = ["14.0924i","5.4887i","14.2337i","7.3637i"]
medCirHdg.ParentStory.Contents = "Healthcare"
medCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

medDataText = myPage.TextFrames.Add()
medDataText.GeometricBounds = ["14.2337i","5.4887i","14.4934i","7.3637i"]
medDataText.ParentStory.Contents = str(med_open_rate) + "% Opens | " + str(med_open_change) + "% Change\n" + str(med_click_rate) + "% Clicks | " + str(med_click_change) + "% Change"
medDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

medArrow = myPage.GraphicLines.Add()
medArrow.GeometricBounds = ["12.5011i","7.6989i","12.7588i","7.6989i"]
medArrow.StrokeColor = colliersDBlue
medArrow.StrokeWeight = 4
medArrow.LeftLineEnd = left_arrow_test(med_perc)
medArrow.RightLineEnd = right_arrow_test(med_perc)

medPercChanged = myPage.TextFrames.Add()
medPercChanged.GeometricBounds = ["12.8552i","7.3551i","13.0986i","8.0426i"]
medPercChanged.ParentStory.Contents = str(med_perc) + "%"
medPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

medComText = myPage.TextFrames.Add()
medComText.GeometricBounds = ["13.0986i","7.3551i","13.6767i","8.0426i"]
medComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
medComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

medComNum = myPage.TextFrames.Add()
medComNum.GeometricBounds = ["13.6767i","7.3551i","13.8199i","8.0426i"]
medComNum.ParentStory.Contents = med_prev_emails
medComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Industrial Emails
indEmailsGraphic = myPage.Place(indEmailsEPS)
indEmailsGraphic = indEmailsGraphic.Item(1)
indEmailsFrame = indEmailsGraphic.Parent
indEmailsFrame.GeometricBounds = ["12.0955i","7.8581i","14.3871i","10.1498i"]
indEmailsGraphic.GeometricBounds = ["12.0955i","7.8581i","14.3871i","10.1498i"]

indCirNum = myPage.TextFrames.Add()
indCirNum.GeometricBounds = ["12.8552i","8.6915i","13.1007i","9.4943i"]
indCirNum.ParentStory.Contents = ind_emails
indCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

indCirText = myPage.TextFrames.Add()
indCirText.GeometricBounds = ["13.0996i","8.6915i","13.5i","9.4943i"]
indCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
indCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

indCirHdg = myPage.TextFrames.Add()
indCirHdg.GeometricBounds = ["14.0924i","8.1554i","14.2337i","10.0304i"]
indCirHdg.ParentStory.Contents = "Industrial"
indCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

indDataText = myPage.TextFrames.Add()
indDataText.GeometricBounds = ["14.2337i","8.1554i","14.4934i","10.0304i"]
indDataText.ParentStory.Contents = str(ind_open_rate) + "% Opens | " + str(ind_open_change) + "% Change\n" + str(ind_click_rate) + "% Clicks | " + str(ind_click_change) + "% Change"
indDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

indArrow = myPage.GraphicLines.Add()
indArrow.GeometricBounds = ["12.5011i","10.3656i","12.7588i","10.3656i"]
indArrow.StrokeColor = colliersDBlue
indArrow.StrokeWeight = 4
indArrow.LeftLineEnd = left_arrow_test(ind_perc)
indArrow.RightLineEnd = right_arrow_test(ind_perc)

indPercChanged = myPage.TextFrames.Add()
indPercChanged.GeometricBounds = ["12.8552i","10.0218i","13.0986i","10.7093i"]
indPercChanged.ParentStory.Contents = str(ind_perc) + "%"
indPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

indComText = myPage.TextFrames.Add()
indComText.GeometricBounds = ["13.0986i","10.0218i","13.6767i","10.7093i"]
indComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
indComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

indComNum = myPage.TextFrames.Add()
indComNum.GeometricBounds = ["13.6767i","10.0218i","13.8199i","10.7093i"]
indComNum.ParentStory.Contents = ind_prev_emails
indComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Investment Emails
invEmailsGraphic = myPage.Place(invEmailsEPS)
invEmailsGraphic = invEmailsGraphic.Item(1)
invEmailsFrame = invEmailsGraphic.Parent
invEmailsFrame.GeometricBounds = ["14.401i","-0.1439i","16.6927i","2.1478i"]
invEmailsGraphic.GeometricBounds = ["14.401i","-0.1439i","16.6927i","2.1478i"]

invCirNum = myPage.TextFrames.Add()
invCirNum.GeometricBounds = ["15.1674i","0.6915i","15.4128i","1.4943i"]
invCirNum.ParentStory.Contents = inv_emails
invCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

invCirText = myPage.TextFrames.Add()
invCirText.GeometricBounds = ["15.4117i","0.6915i","15.84i","1.4943i"]
invCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
invCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

invCirHdg = myPage.TextFrames.Add()
invCirHdg.GeometricBounds = ["16.349i","0.1554i","16.4903i","2.0304i"]
invCirHdg.ParentStory.Contents = "Investment"
invCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

invDataText = myPage.TextFrames.Add()
invDataText.GeometricBounds = ["16.4903i","0.1554i","16.75i","2.0304i"]
invDataText.ParentStory.Contents = str(inv_open_rate) + "% Opens | " + str(inv_open_change) + "% Change\n" + str(inv_click_rate) + "% Clicks | " + str(inv_click_change) + "% Change"
invDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

invArrow = myPage.GraphicLines.Add()
invArrow.GeometricBounds = ["14.8132i","2.3656i","15.071i","2.3656i"]
invArrow.StrokeColor = colliersDBlue
invArrow.StrokeWeight = 4
invArrow.LeftLineEnd = left_arrow_test(inv_perc)
invArrow.RightLineEnd = right_arrow_test(inv_perc)

invPercChanged = myPage.TextFrames.Add()
invPercChanged.GeometricBounds = ["15.1674i","2.0218i","15.4108i","2.7093i"]
invPercChanged.ParentStory.Contents = str(inv_perc) + "%"
invPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

invComText = myPage.TextFrames.Add()
invComText.GeometricBounds = ["15.4108i","2.0218i","15.9889i","2.7093i"]
invComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
invComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

invComNum = myPage.TextFrames.Add()
invComNum.GeometricBounds = ["15.9889i","2.0218i","16.132i","2.7093i"]
invComNum.ParentStory.Contents = inv_prev_emails
invComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Land Emails
landEmailsGraphic = myPage.Place(landEmailsEPS)
landEmailsGraphic = landEmailsGraphic.Item(1)
landEmailsFrame = landEmailsGraphic.Parent
landEmailsFrame.GeometricBounds = ["14.401i","2.5255i","16.6927i","4.8171i"]
landEmailsGraphic.GeometricBounds = ["14.401i","2.5255i","16.6927i","4.8171i"]

landCirNum = myPage.TextFrames.Add()
landCirNum.GeometricBounds = ["15.1674i","3.3582i","15.4128i","4.161i"]
landCirNum.ParentStory.Contents = land_emails
landCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

landCirText = myPage.TextFrames.Add()
landCirText.GeometricBounds = ["15.4117i","3.3582i","15.84i","4.161i"]
landCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
landCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

landCirHdg = myPage.TextFrames.Add()
landCirHdg.GeometricBounds = ["16.349i","2.8221i","16.4903i","4.6971i"]
landCirHdg.ParentStory.Contents = "Land"
landCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

landDataText = myPage.TextFrames.Add()
landDataText.GeometricBounds = ["16.4903i","2.8221i","16.75i","4.6971i"]
landDataText.ParentStory.Contents = str(land_open_rate) + "% Opens | " + str(land_open_change) + "% Change\n" + str(land_click_rate) + "% Clicks | " + str(land_click_change) + "% Change"
landDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

landArrow = myPage.GraphicLines.Add()
landArrow.GeometricBounds = ["14.8132i","5.0322i","15.071i","5.0322i"]
landArrow.StrokeColor = colliersDBlue
landArrow.StrokeWeight = 4
landArrow.LeftLineEnd = left_arrow_test(land_perc)
landArrow.RightLineEnd = right_arrow_test(land_perc)

landPercChanged = myPage.TextFrames.Add()
landPercChanged.GeometricBounds = ["15.1674i","4.6885i","15.4108i","5.376i"]
landPercChanged.ParentStory.Contents = str(land_perc) + "%"
landPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

landComText = myPage.TextFrames.Add()
landComText.GeometricBounds = ["15.4108i","4.6885i","15.9889i","5.376i"]
landComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
landComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

landComNum = myPage.TextFrames.Add()
landComNum.GeometricBounds = ["15.9889i","4.6885i","16.132i","5.376i"]
landComNum.ParentStory.Contents = land_prev_emails
landComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Office Emails
offEmailsGraphic = myPage.Place(offEmailsEPS)
offEmailsGraphic = offEmailsGraphic.Item(1)
offEmailsFrame = offEmailsGraphic.Parent
offEmailsFrame.GeometricBounds = ["14.401i","5.1894i","16.6927i","7.4811i"]
offEmailsGraphic.GeometricBounds = ["14.401i","5.1894i","16.6927i","7.4811i"]

offCirNum = myPage.TextFrames.Add()
offCirNum.GeometricBounds = ["15.1674i","6.0249i","15.4128i","6.8276i"]
offCirNum.ParentStory.Contents = off_emails
offCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

offCirText = myPage.TextFrames.Add()
offCirText.GeometricBounds = ["15.4117i","6.0249i","15.84i","6.8276i"]
offCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
offCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

offCirHdg = myPage.TextFrames.Add()
offCirHdg.GeometricBounds = ["16.349i","5.4887i","16.4903i","7.3637i"]
offCirHdg.ParentStory.Contents = "Office"
offCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

offDataText = myPage.TextFrames.Add()
offDataText.GeometricBounds = ["16.4903i","5.4887i","16.75i","7.3637i"]
offDataText.ParentStory.Contents = str(off_open_rate) + "% Opens | " + str(off_open_change) + "% Change\n" + str(off_click_rate) + "% Clicks | " + str(off_click_change) + "% Change"
offDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

offArrow = myPage.GraphicLines.Add()
offArrow.GeometricBounds = ["14.8132i","7.6989i","15.071i","7.6989i"]
offArrow.StrokeColor = colliersDBlue
offArrow.StrokeWeight = 4
offArrow.LeftLineEnd = left_arrow_test(off_perc)
offArrow.RightLineEnd = right_arrow_test(off_perc)

offPercChanged = myPage.TextFrames.Add()
offPercChanged.GeometricBounds = ["15.1674i","7.3551i","15.4108i","8.0426i"]
offPercChanged.ParentStory.Contents = str(off_perc) + "%"
offPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

offComText = myPage.TextFrames.Add()
offComText.GeometricBounds = ["15.4108i","7.3551i","15.9889i","8.0426i"]
offComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
offComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

offComNum = myPage.TextFrames.Add()
offComNum.GeometricBounds = ["15.9889i","7.3551i","16.132i","8.0426i"]
offComNum.ParentStory.Contents = off_prev_emails
offComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Retail Emails
retEmailsGraphic = myPage.Place(retEmailsEPS)
retEmailsGraphic = retEmailsGraphic.Item(1)
retEmailsFrame = retEmailsGraphic.Parent
retEmailsFrame.GeometricBounds = ["14.401i","7.8581i","16.6927i","10.1498i"]
retEmailsGraphic.GeometricBounds = ["14.401i","7.8581i","16.6927i","10.1498i"]

retCirNum = myPage.TextFrames.Add()
retCirNum.GeometricBounds = ["15.1674i","8.6915i","15.4128i","9.4943i"]
retCirNum.ParentStory.Contents = ret_emails
retCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

retCirText = myPage.TextFrames.Add()
retCirText.GeometricBounds = ["15.4117i","8.6915i","15.84i","9.4943i"]
retCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
retCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

retCirHdg = myPage.TextFrames.Add()
retCirHdg.GeometricBounds = ["16.349i","8.1554i","16.4903i","10.0304i"]
retCirHdg.ParentStory.Contents = "Retail"
retCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

retDataText = myPage.TextFrames.Add()
retDataText.GeometricBounds = ["16.4903i","8.1554i","16.75i","10.0304i"]
retDataText.ParentStory.Contents = str(ret_open_rate) + "% Opens | " + str(ret_open_change) + "% Change\n" + str(ret_click_rate) + "% Clicks | " + str(ret_click_change) + "% Change"
retDataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

retArrow = myPage.GraphicLines.Add()
retArrow.GeometricBounds = ["14.8132i","10.3656i","15.071i","10.3656i"]
retArrow.StrokeColor = colliersDBlue
retArrow.StrokeWeight = 4
retArrow.LeftLineEnd = left_arrow_test(ret_perc)
retArrow.RightLineEnd = right_arrow_test(ret_perc)

retPercChanged = myPage.TextFrames.Add()
retPercChanged.GeometricBounds = ["15.1674i","10.0218i","15.4108i","10.7093i"]
retPercChanged.ParentStory.Contents = str(ret_perc) + "%"
retPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

retComText = myPage.TextFrames.Add()
retComText.GeometricBounds = ["15.4108i","10.0218i","15.9889i","10.7093i"]
retComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
retComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

retComNum = myPage.TextFrames.Add()
retComNum.GeometricBounds = ["15.9889i","10.0218i","16.132i","10.7093i"]
retComNum.ParentStory.Contents = ret_prev_emails
retComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle
