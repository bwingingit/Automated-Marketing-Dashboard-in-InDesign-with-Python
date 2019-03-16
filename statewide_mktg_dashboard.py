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

# SETUP
# Pull in all the data we will need
inputs_csv = 'marketing-dashboard-inputs2.csv'
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
co_colors = (\
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
special_campaign_type = lambda c: c if (c in property_campaigns_list) or ('Internal Newsletters' in str(c)) or ('Market Reports' in str(c)) else 'Custom Special Campaigns'
open_rate_calc = lambda row: round((row.Opened/row.Sent*100),1) if row.Sent > 1 else 0
click_rate_calc = lambda row: round((row.Clicked/row.Sent*100),2) if row.Sent > 1 else 0
sent_perc_change = lambda row: int((row.Sent - row.Prev_Sent)/(row.Prev_Sent)*100) if row.Prev_Sent > 0 else 100
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

def df_to_list(df):
    df_values = [y for x in df.values.tolist() for y in x]
    return df.columns.tolist() + df_values

def blue_circle(center,radius):
    return Circle(center,radius,color="#00467f",alpha=1)
def white_circle(center,radius):
    return Circle(center,radius,color="#ffffff",alpha=1)
def perc_change_wedge(center,radius,perc):
    return Wedge(center, radius, inner_theta1(perc), inner_theta2(perc), color=patch_color(perc), alpha=inner_alpha(perc))
def outer_white_circle(radius,perc):
    return Circle((.5,.5), radius, color="#ffffff", alpha=outline_alpha(perc))
def outer_perc_change_wedge(radius,perc):
    return Wedge((.5,.5), radius, outer_theta1(perc), outer_theta2(perc), color=patch_color(perc), alpha=outer_alpha(perc))

def hs_page_type(p):
    if 'sign-up' in p:
        return 'Sign-Up'
    elif 'subscriptions' in p:
        return 'Sign-Up'
    elif 'report' in p:
        return 'Research'
    else:
        return 'Landing'
      
def proposal_outcome(row):
    if row.Status == 'WIN':
        return 'Won'
    elif row.Status == 'LOSS':
        return 'Lost'
    elif row.Status == "PULLED":
        return 'Pulled'
    else:
        return 'Outstanding'

def outcomes_output(value):
  try:
    return grouped_proposals.loc[grouped_proposals['Outcome'] == value, 'Proposals'].item()
  except IndexError:
    return 0
  except ValueError:
    return 0

# Get the file links for the social icons
twitterEPS = cwd + '\\twitter.eps'
facebookEPS = cwd + '\\facebook.eps'
instagramEPS = cwd + '\\instagram.eps'

# Do some quick early calculations to get our recent month and previous month variables
emails_csv = inputs_df.iloc[0]['CSV File Name or Numbers']
while '.csv' not in emails_csv:
    emails_csv = input('Please provide the name of the CSV file, including the .csv, for the two months of email data being compared: ')
emails_df = pd.read_csv(emails_csv)
emails_df = emails_df[emails_df.Sent > 1]
emails_df = emails_df.rename(columns={'Send Date (Your time zone)':'Send Date'})
emails_df['Send Date'] = pd.to_datetime(emails_df['Send Date'])
emails_df['Send Month'] = emails_df['Send Date'].dt.month
emails_df['Send Year'] = emails_df['Send Date'].dt.year
if emails_df['Send Date'].max().month == 1:
    rec_month = 1
    prev_month = 12
else:
    rec_month = emails_df['Send Date'].max().month
    prev_month = rec_month - 1
recent_month_name = month_number[rec_month]
recent_mo = mo_num[rec_month]
prev_mo = mo_num[prev_month]
recent_yr = emails_df['Send Date'].max().year

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
strokeNone = myDocument.Swatches.Item("None")
myArrowHead = 1937203560
noArrowHead = 1852796517
myHeaderRow = 1162375799
myVAlignCenter = 1667591796
myVAlignBottom = 1651471469
myGraphicCellType = 1701728329
myEmptyFitProp = 1718185072

# set up colors
# Dark Blue
coDBlue = myDocument.Colors.Add()
coDBlue.Name = "CO Dark Blue"
coDBlue.ColorValue = [100, 57, 0, 38]
# Light Blue
coLBlue = myDocument.Colors.Add()
coLBlue.Name = "CO Light Blue"
coLBlue.ColorValue = [100, 10, 0, 10]
# 80% Gray
co80Gray = myDocument.Colors.Add()
co80Gray.Name = "80 Gray"
co80Gray.ColorValue = [0, 0, 0, 80]
# Yellow
coYellow = myDocument.Colors.Add()
coYellow.Name = "CO Yellow"
coYellow.ColorValue = [0, 24, 94, 0]
# Red
coRed = myDocument.Colors.Add()
coRed.Name = "CO Red"
coRed.ColorValue = [0, 95, 100, 0]
# Dark Red
coDRed = myDocument.Colors.Add()
coDRed.Name = "CO Dark Red"
coDRed.ColorValue = [0, 95, 100, 29]
# White
coWhite = myDocument.Colors.Add()
coWhite.Name = "CO White"
coWhite.ColorValue = [0, 0, 0, 0]

# Create a page
myPage = myDocument.Pages.Item(1)

# set up paragraph styles
# fix [Basic Paragraph]
basicParagraph = myDocument.ParagraphStyles.Item("[Basic Paragraph]")
basicParagraph.FillColor = co80Gray
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
subheadingStyle.FillColor = coWhite

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
tableHeadingCells.BottomEdgeStrokeColor = co80Gray
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
basicCells.BottomEdgeStrokeColor = co80Gray
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
circleNumbersStyle.FillColor = coWhite

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
circleTextStyle.FillColor = coWhite

# Percent Changed
percentChangedStyle = myDocument.ParagraphStyles.Add()
percentChangedStyle.Name = "Percent_Changed"
percentChangedStyle.BasedOn = basicParagraph
percentChangedStyle.PointSize = 16
percentChangedStyle.FontStyle = "Ultra"
percentChangedStyle.Justification = myCenterAlign
percentChangedStyle.FillColor = coDBlue

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
negativeDataTextStyle.FillColor = coDRed

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

# Adding in a dictionary of geometric bounds
# covers most of these
# other geometric bounds used in loops will be included in those dictionaries
# closer to the loops themselves
geo_bounds_misc = {}
geo_bounds_misc['myDshbdTitle'] = ["0.342i", "0.25i", "0.6515i", "3.7095i"]
geo_bounds_misc['myDshbdMonth'] = ["0.6515i", "0.25i", "0.9462i", "3.7095i"]
geo_bounds_misc['socialRect'] = ["1.084i","0.25i","1.4211i","5.5i"]
geo_bounds_misc['socialSubH'] = ["1.1944i","1.3477i","1.335i","4.4023i"]
geo_bounds_misc['webRect'] = ["0.25i","5.5i","0.5871i","10.75i"]
geo_bounds_misc['webSubH'] = ["0.3604i","6.5977i","0.5i","9.6523i"]
geo_bounds_misc['mktgRect'] = ["5.9659i","0.25i","6.302i","10.75i"]
geo_bounds_misc['mktgSubH'] = ["6.0753i","3.9727i","6.208i","7.0272i"]
geo_bounds_misc['emailsRect'] = ["8.7025i","0.25i","9.0396i","5.5i"]
geo_bounds_misc['emailsSubH'] = ["8.8129i","1.3477i","8.96i","4.4023i"]
geo_bounds_misc['socialFrame'] = ["1.2829i","-0.0883i","3.5746i","2.2033i"]
geo_bounds_misc['socialCirNum'] = ["2.1148i","0.7405i","2.3602i","1.5432i"]
geo_bounds_misc['socialCirText'] = ["2.3591i","0.7405i","2.6182i","1.5432i"]
geo_bounds_misc['socialArrow'] = ["1.7i","2.4145i","1.9578i","2.4145i"]
geo_bounds_misc['socialPercChanged'] = ["2.0542i","2.0995i","2.2976i","2.7295i"]
geo_bounds_misc['socialComText'] = ["2.2976i","2.0995i","2.7715i","2.7295i"]
geo_bounds_misc['socialComNum'] = ["2.7715i","2.0995i","2.9147i","2.7295i"]
geo_bounds_misc['twitterIconFrame'] = ["1.841i","3.3166i","2.1755i","3.6511i"]
geo_bounds_misc['twitterData'] = ["1.8703i","3.8065i","2.1339i","5.4769i"]
geo_bounds_misc['facebookIconFrame'] = ["2.3191i","3.3166i","2.6536i","3.6511i"]
geo_bounds_misc['facebookData'] = ["2.3546i","3.8065i","2.6182i","5.4769i"]
geo_bounds_misc['instagramIconFrame'] = ["2.7715i","3.3166i","3.106i","3.6511i"]
geo_bounds_misc['instagramData'] = ["2.8021i","3.8065i","3.0657i","5.4769i"]
geo_bounds_misc['socialTableText'] = ["3.3735i","0.26i","5.7768i","5.6i"]
geo_bounds_misc['websiteFrame'] = ["0.5422i","5.3479i","3.6394i","8.4451i"]
geo_bounds_misc['websiteCirNum'] = ["1.7259i","6.3318i","2.0021i","7.489i"]
geo_bounds_misc['websiteCirText'] = ["2.0846i","6.3318i","2.4864i","7.489i"]
geo_bounds_misc['campaignTableText'] = ["0.7762i","8.4498i","4.0896i","10.74i"]
geo_bounds_misc['pagesTableText'] = ["3.5466i","5.7598i","5.7768i","8.1096i"]
geo_bounds_misc['propSubText'] = ["4.287i","8.6625i","4.6215i","10.5i"]
geo_bounds_misc['propSubHdg'] = ["4.62i","8.6625i","4.9571i","10.5i"]
geo_bounds_misc['resSubText'] = ["5.1066i","8.6625i","5.4412i","10.5i"]
geo_bounds_misc['resSubHdg'] = ["5.4397i","8.6625i","5.7768i","10.5i"]
geo_bounds_misc['opensText'] = ["9.1702i","0.2292i","9.5047i","1.3234i"]
geo_bounds_misc['opensHdg'] = ["9.5033i","0.2292i","9.66i","1.3234i"]
geo_bounds_misc['openRateText'] = ["9.1702i","1.6133i","9.5047i","2.7075i"]
geo_bounds_misc['openRateHdg'] = ["9.5033i","1.6133i","9.66i","2.7075i"]
geo_bounds_misc['clicksText'] = ["9.1702i","2.9974i","9.5047i","4.0916i"]
geo_bounds_misc['clicksHdg'] = ["9.5033i","2.9974i","9.66i","4.0916i"]
geo_bounds_misc['clickRateText'] = ["9.1702i","4.3814i","9.5047i","5.4757i"]
geo_bounds_misc['clickRateHdg'] = ["9.5033i","4.3814i","9.66i","5.4757i"]
geo_bounds_misc['emailsTableText'] = ["9.894i","0.25i","12.1819i","6.107i"]
geo_bounds_misc['companyEmailsFrame'] = ["8.45i","6.1458i","12.45i","10.3542i"]
geo_bounds_misc['companyEmailsCirNum'] = ["10.0912i","7.6768i","10.4131i","8.834i"]
geo_bounds_misc['companyEmailsCirText'] = ["10.4499i","7.8581i","10.8795i","8.6527i"]
geo_bounds_misc['companyEmailsArrow'] = ["9.5926i","10.3334i","9.9896i","10.3334i"]
geo_bounds_misc['companyEmailsPercChanged'] = ["10.0912i","9.9169i","10.4131i","10.75i"]
geo_bounds_misc['companyEmailsComText'] = ["10.4131i","9.9169i","10.8795i","10.75i"]
geo_bounds_misc['companyEmailsComNum'] = ["10.8795i","9.9169i","11.0689i","10.75i"]
geo_bounds_misc['adCirHdg'] = ["8.2337i","5.5412i","8.375i","7.3112i"]
geo_bounds_misc['adFrame'] = ["6.2222i","5.1894i","8.5139i","7.4811i"]
geo_bounds_misc['adCirNum'] = ["7.0571i","6.0249i","7.3025i","6.8276i"]
geo_bounds_misc['adCirText'] = ["7.315i","6.0249i","7.5659i","6.8276i"]
geo_bounds_misc['adPercChanged'] = ["6.9965i","7.3839i","7.2399i","8.0139i"]
geo_bounds_misc['adComText'] = ["7.2399i","7.2938i","7.7138i","8.104i"]
geo_bounds_misc['adComNum'] = ["7.7138i","7.3839i","7.857i","8.0139i"]
geo_bounds_misc['propCirHdg'] = ["8.2337i","8.2079i","8.375i","9.9779i"]
geo_bounds_misc['propFrame'] = ["6.2219i","8.00068i","8.3889i","10.1738i"]
geo_bounds_misc['propCirNum'] = ["7.0571i","8.6915i","7.3025i","9.4943i"]
geo_bounds_misc['propCirText'] = ["7.315i","8.6915i","7.5659i","9.4943i"]
geo_bounds_misc['propPercChanged'] = ["6.9965i","10.0505i","7.2399i","10.6806i"]
geo_bounds_misc['propComText'] = ["7.2399i","10.0505i","7.7138i","10.6806i"]
geo_bounds_misc['propComNum'] = ["7.7138i","10.0505i","7.857i","10.6806i"]

myDshbdTitle = myPage.TextFrames.Add()
myDshbdTitle.GeometricBounds = geo_bounds_misc['myDshbdTitle']
myDshbdTitle.ParentStory.Contents = "Marketing Dashboard"
myDshbdTitle.ParentStory.Characters.Item(1).appliedParagraphStyle = heading1Style

myDshbdMonth = myPage.TextFrames.Add()
myDshbdMonth.GeometricBounds = geo_bounds_misc['myDshbdMonth']
myDshbdMonth.ParentStory.Contents = "South Carolina | " + recent_month_name + str(recent_yr)
myDshbdMonth.ParentStory.Characters.Item(1).appliedParagraphStyle = heading2Style

# Now for the email calculations and layout
# first, the email bar
emailsRect = myPage.Rectangles.Add()
emailsRect.GeometricBounds = geo_bounds_misc['emailsRect']
emailsRect.FillColor = coDBlue
emailsRect.StrokeColor = strokeNone
emailsSubH = myPage.TextFrames.Add()
emailsSubH.GeometricBounds = geo_bounds_misc['emailsSubH']
emailsSubH.ParentStory.Contents = "Company-wide Emails"
emailsSubH.ParentStory.Characters.Item(1).appliedParagraphStyle = subheadingStyle

# now the calcs for opens and clicks
recent_emails = emails_df[emails_df['Send Month'] == rec_month]
prev_emails = emails_df[emails_df['Send Month'] == prev_month]

total_emails = recent_emails['Sent'].sum()
total_opens = recent_emails['Opened'].sum()
total_open_rate = round(((total_opens/total_emails)*100),2)
total_clicks = recent_emails['Clicked'].sum()
total_click_rate = round(((total_clicks/total_emails)*100),2)

opensText = myPage.TextFrames.Add()
opensText.GeometricBounds = geo_bounds_misc['opensText']
opensText.ParentStory.Contents = '{:,}'.format(total_opens)
opensText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
opensHdg = myPage.TextFrames.Add()
opensHdg.GeometricBounds = geo_bounds_misc['opensHdg']
opensHdg.ParentStory.Contents = "Opens"
opensHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

openRateText = myPage.TextFrames.Add()
openRateText.GeometricBounds = geo_bounds_misc['openRateText']
openRateText.ParentStory.Contents = str(total_open_rate) + "%"
openRateText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
openRateHdg = myPage.TextFrames.Add()
openRateHdg.GeometricBounds = geo_bounds_misc['openRateHdg']
openRateHdg.ParentStory.Contents = "Open Rate"
openRateHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

clicksText = myPage.TextFrames.Add()
clicksText.GeometricBounds = geo_bounds_misc['clicksText']
clicksText.ParentStory.Contents = '{:,}'.format(total_clicks)
clicksText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
clicksHdg = myPage.TextFrames.Add()
clicksHdg.GeometricBounds = geo_bounds_misc['clicksHdg']
clicksHdg.ParentStory.Contents = "Clicks"
clicksHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

clickRateText = myPage.TextFrames.Add()
clickRateText.GeometricBounds = geo_bounds_misc['clickRateText']
clickRateText.ParentStory.Contents = str(total_click_rate) + "%"
clickRateText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
clickRateHdg = myPage.TextFrames.Add()
clickRateHdg.GeometricBounds = geo_bounds_misc['clickRateHdg']
clickRateHdg.ParentStory.Contents = "Click Rate"
clickRateHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

# now for the email recipients by campaign calcs and chart
combined_emails = recent_emails[['Campaign','Sent']].reset_index(drop=True)
combined_emails['Long_Email_Campaign'] = combined_emails.Campaign.apply(special_campaign_type)
combined_emails['Email_Campaign'] = combined_emails.Long_Email_Campaign.apply(lambda x: x.split()[1])
grouped_combined_emails = combined_emails.groupby('Email_Campaign').Sent.sum().reset_index()
total_emails_sent = grouped_combined_emails.Sent.sum()
grouped_combined_emails['Perc_Total'] = grouped_combined_emails.Sent.apply(lambda s: int((s/total_emails_sent)*100))
grouped_combined_emails['Label'] = grouped_combined_emails.apply(lambda row: str(row.Perc_Total) + "%\n" + row.Email_Campaign, axis=1)
grouped_combined_emails = grouped_combined_emails.sort_values(by=['Email_Campaign'],ascending=True).reset_index(drop=True)
figEm, axEm = plt.subplots(figsize=(4.2, 4), subplot_kw=dict(aspect="equal"))
wedgesEm,textsEm = axEm.pie(grouped_combined_emails['Perc_Total'].tolist(),
                        colors=co_colors,
                        labels=grouped_combined_emails['Label'].tolist(),
                        labeldistance=1.2,
                        startangle=90,
                        wedgeprops=dict(width=0.25,linewidth=1,edgecolor='w'))
for t in textsEm:
    t.set_horizontalalignment('center')
total_prev_emails = prev_emails['Sent'].sum()
sent_perc = int((total_emails - total_prev_emails)/(total_prev_emails)*100)
axEm.add_artist(perc_change_wedge((0,0),0.7,sent_perc))
axEm.add_artist(white_circle((0,0),0.55))
axEm.add_artist(blue_circle((0,0),0.525))
plt.tight_layout()
figEm.savefig('total_emails_by_campaign.eps',transparent=True)
totalEmailsEPS = cwd + '\\total_emails_by_campaign.eps'

companyEmailsFrame = myPage.Rectangles.Add()
companyEmailsFrame.GeometricBounds = geo_bounds_misc['companyEmailsFrame']
companyEmailsFrame.StrokeColor = strokeNone
companyEmailsFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
companyEmailsGraphic = companyEmailsFrame.Place(totalEmailsEPS)

companyEmailsCirNum = myPage.TextFrames.Add()
companyEmailsCirNum.GeometricBounds = geo_bounds_misc['companyEmailsCirNum']
companyEmailsCirNum.ParentStory.Contents = '{:,}'.format(total_emails)
companyEmailsCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = largeCircleNumbersStyle

companyEmailsCirText = myPage.TextFrames.Add()
companyEmailsCirText.GeometricBounds = geo_bounds_misc['companyEmailsCirText']
companyEmailsCirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
companyEmailsCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

companyEmailsArrow = myPage.GraphicLines.Add()
companyEmailsArrow.GeometricBounds = geo_bounds_misc['companyEmailsArrow']
companyEmailsArrow.StrokeColor = coDBlue
companyEmailsArrow.StrokeWeight = 7
companyEmailsArrow.LeftLineEnd = left_arrow_test(sent_perc)
companyEmailsArrow.RightLineEnd = right_arrow_test(sent_perc)

companyEmailsPercChanged = myPage.TextFrames.Add()
companyEmailsPercChanged.GeometricBounds = geo_bounds_misc['companyEmailsPercChanged']
companyEmailsPercChanged.ParentStory.Contents = str(sent_perc) + "%"
companyEmailsPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = largePercentChangeStyle

companyEmailsComText = myPage.TextFrames.Add()
companyEmailsComText.GeometricBounds = geo_bounds_misc['companyEmailsComText']
companyEmailsComText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
companyEmailsComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

companyEmailsComNum = myPage.TextFrames.Add()
companyEmailsComNum.GeometricBounds = geo_bounds_misc['companyEmailsComNum']
companyEmailsComNum.ParentStory.Contents = '{:,}'.format(total_prev_emails)
companyEmailsComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Now for the top 5 performing emails
recent_property_emails = recent_emails[['Email Name','Campaign','Sent','Open Rate','Click Rate','Unsubscribed','Hard Bounced','Soft Bounced']].reset_index(drop=True)
recent_property_emails = recent_property_emails[recent_property_emails.Campaign.isin(property_campaigns_list)]
recent_property_emails = recent_property_emails.sort_values(by=['Click Rate'],ascending=False).reset_index(drop=True)
top_property_emails = recent_property_emails.iloc[:5].reset_index()
top_property_emails['Product'] = top_property_emails.Campaign.apply(lambda x: x.split()[1])
top_property_emails['Open_Rate'] = pd.to_numeric(top_property_emails['Open Rate']).round(1).apply(lambda x: str(x) + '%')
top_property_emails['Click_Rate'] = pd.to_numeric(top_property_emails['Click Rate']).round(1).apply(lambda x: str(x) + '%')
top_property_emails['Bounce'] = (top_property_emails['Hard Bounced'] + top_property_emails['Soft Bounced']).apply(lambda x: str(x))
top_property_emails['Sent'] = top_property_emails['Sent'].apply(lambda x: '{:,}'.format(x))
top_property_emails['Unsubscribed'] = top_property_emails['Unsubscribed'].apply(lambda x: str(x))
top_emails = top_property_emails[['Email Name','Product','Sent','Open_Rate','Click_Rate','Unsubscribed','Bounce']].rename(columns={'Open_Rate':'Open Rate','Click_Rate':'Click Rate'})
top_emails_list = df_to_list(top_emails)

emailsTableText = myPage.TextFrames.Add()
emailsTableText.GeometricBounds = geo_bounds_misc['emailsTableText']
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
emailsTable.Columns.Item(5).FillColor = coYellow
emailsTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
emailsTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
emailsTable.Rows.Item(1).VerticalJustification = myVAlignBottom
for i in range(2,7):
    emailsTable.Rows.Item(i).VerticalJustification = myVAlignCenter
emailsTable.Contents = top_emails_list
emailsrangetotal = emailscolumncount * emailsrowcount
for i in range(1,emailsrangetotal,emailscolumncount):
    emailsTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign

# Now for the first loop
# graphics and data by producty type
# let's start with some preliminary calcs
campaigns_columns = ['Campaign','Sent','Opened','Clicked']
recent_campaigns = recent_emails[campaigns_columns].reset_index(drop=True)
recent_campaigns = recent_campaigns[recent_campaigns.Campaign.isin(property_campaigns_list)]
recent_campaigns_group = recent_campaigns.groupby('Campaign')['Sent','Opened','Clicked'].sum().reset_index()
prev_campaigns = prev_emails[campaigns_columns].reset_index(drop=True)
prev_campaigns.rename(columns={'Sent':'Prev_Sent','Opened':'Prev_Opened','Clicked':'Prev_Clicked'},inplace=True)
prev_campaigns = prev_campaigns[prev_campaigns.Campaign.isin(property_campaigns_list)]
prev_campaigns_group = prev_campaigns.groupby('Campaign')['Prev_Sent','Prev_Opened','Prev_Clicked'].sum().reset_index()
campaigns_comparison = pd.merge(recent_campaigns_group,prev_campaigns_group,how='outer')
campaigns_comparison['Product'] = campaigns_comparison.Campaign.apply(lambda x: x.split()[1])
campaign_totals = campaigns_comparison.groupby('Product')['Sent','Prev_Sent','Opened','Prev_Opened','Clicked','Prev_Clicked'].sum().reset_index()
campaign_totals['Open_Rate'] = campaign_totals.apply(open_rate_calc,axis=1)
campaign_totals['Click_Rate'] = campaign_totals.apply(click_rate_calc,axis=1)
campaign_totals['Sent_Change'] = campaign_totals.apply(sent_perc_change,axis=1)
campaign_totals['Open_Change'] = campaign_totals.apply(open_perc_change,axis=1)
campaign_totals['Click_Change'] = campaign_totals.apply(click_perc_change,axis=1)
campaign_totals['Total_Emails'] = pd.to_numeric(campaign_totals.Sent)
campaign_totals['Prev_Total'] = pd.to_numeric(campaign_totals.Prev_Sent)
campaign_totals_display = campaign_totals[['Product','Total_Emails','Sent_Change','Prev_Total','Open_Rate','Open_Change','Click_Rate','Click_Change']]

# now we'll add dictionary of variable and geometric bounds
graphics = {}
graphics['development'] = {}
graphics['development']['name'] = 'Development'
graphics['development']['str_filter'] = 'Development'
graphics['development']['eps_file'] = 'development-emails.eps'
graphics['development']['emailsFrame'] = ["12.0955i","-0.1439i","14.3871i","2.1478i"]
graphics['development']['cirNum'] = ["12.8552i","0.6915i","13.1007i","1.4943i"]
graphics['development']['cirText'] = ["13.0996i","0.6915i","13.5i","1.4943i"]
graphics['development']['cirHdg'] = ["14.0924i","0.1554i","14.2337i","2.0304i"]
graphics['development']['dataText'] = ["14.2337i","0.1554i","14.4934i","2.0304i"]
graphics['development']['arrow'] = ["12.5011i","2.3656i","12.7588i","2.3656i"]
graphics['development']['percChanged'] = ["12.8552i","2.0218i","13.0986i","2.7093i"]
graphics['development']['comText'] = ["13.0986i","2.0218i","13.6767i","2.7093i"]
graphics['development']['comNum'] = ["13.6767i","2.0218i","13.8199i","2.7093i"]
graphics['flex'] = {}
graphics['flex']['name'] = 'Flex'
graphics['flex']['str_filter'] = 'Flex'
graphics['flex']['eps_file'] = 'flex-emails.eps'
graphics['flex']['emailsFrame'] = ["12.0955i","2.5255i","14.3871i","4.8171i"]
graphics['flex']['cirNum'] = ["12.8552i","3.3582i","13.1007i","4.161i"]
graphics['flex']['cirText'] = ["13.0996i","3.3582i","13.5i","4.161i"]
graphics['flex']['cirHdg'] = ["14.0924i","2.8221i","14.2337i","4.6971i"]
graphics['flex']['dataText'] = ["14.2337i","2.8221i","14.4934i","4.6971i"]
graphics['flex']['arrow'] = ["12.5011i","5.0322i","12.7588i","5.0322i"]
graphics['flex']['percChanged'] = ["12.8552i","4.6885i","13.0986i","5.376i"]
graphics['flex']['comText'] = ["13.0986i","4.6885i","13.6767i","5.376i"]
graphics['flex']['comNum'] = ["13.6767i","4.6885i","13.8199i","5.376i"]
graphics['healthcare'] = {}
graphics['healthcare']['name'] = 'Healthcare'
graphics['healthcare']['str_filter'] = 'Medical'
graphics['healthcare']['eps_file'] = 'medical-emails.eps'
graphics['healthcare']['emailsFrame'] = ["12.0955i","5.1894i","14.3871i","7.4811i"]
graphics['healthcare']['cirNum'] = ["12.8552i","6.0249i","13.1007i","6.8276i"]
graphics['healthcare']['cirText'] = ["13.0996i","6.0249i","13.5i","6.8276i"]
graphics['healthcare']['cirHdg'] = ["14.0924i","5.4887i","14.2337i","7.3637i"]
graphics['healthcare']['dataText'] = ["14.2337i","5.4887i","14.4934i","7.3637i"]
graphics['healthcare']['arrow'] = ["12.5011i","7.6989i","12.7588i","7.6989i"]
graphics['healthcare']['percChanged'] = ["12.8552i","7.3551i","13.0986i","8.0426i"]
graphics['healthcare']['comText'] = ["13.0986i","7.3551i","13.6767i","8.0426i"]
graphics['healthcare']['comNum'] = ["13.6767i","7.3551i","13.8199i","8.0426i"]
graphics['industrial'] = {}
graphics['industrial']['name'] = 'Industrial'
graphics['industrial']['str_filter'] = 'Industrial'
graphics['industrial']['eps_file'] = 'industrial-emails.eps'
graphics['industrial']['emailsFrame'] = ["12.0955i","7.8581i","14.3871i","10.1498i"]
graphics['industrial']['cirNum'] = ["12.8552i","8.6915i","13.1007i","9.4943i"]
graphics['industrial']['cirText'] = ["13.0996i","8.6915i","13.5i","9.4943i"]
graphics['industrial']['cirHdg'] = ["14.0924i","8.1554i","14.2337i","10.0304i"]
graphics['industrial']['dataText'] = ["14.2337i","8.1554i","14.4934i","10.0304i"]
graphics['industrial']['arrow'] = ["12.5011i","10.3656i","12.7588i","10.3656i"]
graphics['industrial']['percChanged'] = ["12.8552i","10.0218i","13.0986i","10.7093i"]
graphics['industrial']['comText'] = ["13.0986i","10.0218i","13.6767i","10.7093i"]
graphics['industrial']['comNum'] = ["13.6767i","10.0218i","13.8199i","10.7093i"]
graphics['investment'] = {}
graphics['investment']['name'] = 'Investment'
graphics['investment']['str_filter'] = 'Investment'
graphics['investment']['eps_file'] = 'investment-emails.eps'
graphics['investment']['emailsFrame'] = ["14.401i","-0.1439i","16.6927i","2.1478i"]
graphics['investment']['cirNum'] = ["15.1674i","0.6915i","15.4128i","1.4943i"]
graphics['investment']['cirText'] = ["15.4117i","0.6915i","15.84i","1.4943i"]
graphics['investment']['cirHdg'] = ["16.349i","0.1554i","16.4903i","2.0304i"]
graphics['investment']['dataText'] = ["16.4903i","0.1554i","16.75i","2.0304i"]
graphics['investment']['arrow'] = ["14.8132i","2.3656i","15.071i","2.3656i"]
graphics['investment']['percChanged'] = ["15.1674i","2.0218i","15.4108i","2.7093i"]
graphics['investment']['comText'] = ["15.4108i","2.0218i","15.9889i","2.7093i"]
graphics['investment']['comNum'] = ["15.9889i","2.0218i","16.132i","2.7093i"]
graphics['land'] = {}
graphics['land']['name'] = 'Land'
graphics['land']['str_filter'] = 'Land'
graphics['land']['eps_file'] = 'land-emails.eps'
graphics['land']['emailsFrame'] = ["14.401i","2.5255i","16.6927i","4.8171i"]
graphics['land']['cirNum'] = ["15.1674i","3.3582i","15.4128i","4.161i"]
graphics['land']['cirText'] = ["15.4117i","3.3582i","15.84i","4.161i"]
graphics['land']['cirHdg'] = ["16.349i","2.8221i","16.4903i","4.6971i"]
graphics['land']['dataText'] = ["16.4903i","2.8221i","16.75i","4.6971i"]
graphics['land']['arrow'] = ["14.8132i","5.0322i","15.071i","5.0322i"]
graphics['land']['percChanged'] = ["15.1674i","4.6885i","15.4108i","5.376i"]
graphics['land']['comText'] = ["15.4108i","4.6885i","15.9889i","5.376i"]
graphics['land']['comNum'] = ["15.9889i","4.6885i","16.132i","5.376i"]
graphics['office'] = {}
graphics['office']['name'] = 'Office'
graphics['office']['str_filter'] = 'Office'
graphics['office']['eps_file'] = 'office-emails.eps'
graphics['office']['emailsFrame'] = ["14.401i","5.1894i","16.6927i","7.4811i"]
graphics['office']['cirNum'] = ["15.1674i","6.0249i","15.4128i","6.8276i"]
graphics['office']['cirText'] = ["15.4117i","6.0249i","15.84i","6.8276i"]
graphics['office']['cirHdg'] = ["16.349i","5.4887i","16.4903i","7.3637i"]
graphics['office']['dataText'] = ["16.4903i","5.4887i","16.75i","7.3637i"]
graphics['office']['arrow'] = ["14.8132i","7.6989i","15.071i","7.6989i"]
graphics['office']['percChanged'] = ["15.1674i","7.3551i","15.4108i","8.0426i"]
graphics['office']['comText'] = ["15.4108i","7.3551i","15.9889i","8.0426i"]
graphics['office']['comNum'] = ["15.9889i","7.3551i","16.132i","8.0426i"]
graphics['retail'] = {}
graphics['retail']['name'] = 'Retail'
graphics['retail']['str_filter'] = 'Retail'
graphics['retail']['eps_file'] = 'retail-emails.eps'
graphics['retail']['emailsFrame'] = ["14.401i","7.8581i","16.6927i","10.1498i"]
graphics['retail']['cirNum'] = ["15.1674i","8.6915i","15.4128i","9.4943i"]
graphics['retail']['cirText'] = ["15.4117i","8.6915i","15.84i","9.4943i"]
graphics['retail']['cirHdg'] = ["16.349i","8.1554i","16.4903i","10.0304i"]
graphics['retail']['dataText'] = ["16.4903i","8.1554i","16.75i","10.0304i"]
graphics['retail']['arrow'] = ["14.8132i","10.3656i","15.071i","10.3656i"]
graphics['retail']['percChanged'] = ["15.1674i","10.0218i","15.4108i","10.7093i"]
graphics['retail']['comText'] = ["15.4108i","10.0218i","15.9889i","10.7093i"]
graphics['retail']['comNum'] = ["15.9889i","10.0218i","16.132i","10.7093i"]

# and the loop
for graphic in graphics.values():
  try:
    totals = campaign_totals_display[campaign_totals_display.Product == graphic['str_filter']].reset_index(drop=True)
    emails = totals.iloc[0]['Total_Emails']
    perc = totals.iloc[0]['Sent_Change']
    prev_emails = totals.iloc[0]['Prev_Total']
    open_rate = totals.iloc[0]['Open_Rate']
    open_change = totals.iloc[0]['Open_Change']
    click_rate = totals.iloc[0]['Click_Rate']
    click_change = totals.iloc[0]['Click_Change']
  except IndexError:
    emails = 0
    perc = 0
    prev_emails = 0
    open_rate = 0
    open_change = 0
    click_rate = 0
    click_change = 0
  fig,ax = plt.subplots(figsize=(2.3,2.3),subplot_kw=dict(aspect='equal'))
  ax.add_artist(outer_perc_change_wedge(0.5,perc))
  ax.add_artist(outer_white_circle(0.435,perc))
  ax.add_artist(perc_change_wedge((0.5,0.5),0.425,perc))
  ax.add_artist(white_circle((0.5,0.5),0.33))
  ax.add_artist(blue_circle((0.5,0.5),0.32))
  plt.axis('off')
  plt.tight_layout()
  fig.savefig(graphic['eps_file'],transparent=True)
  eps_file_path = cwd + '\\' + graphic['eps_file']
  # now to place all of this on the page
  # start with the graphic
  emailsFrame = myPage.Rectangles.Add()
  emailsFrame.GeometricBounds = graphic['emailsFrame']
  emailsFrame.StrokeColor = strokeNone
  emailsFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
  emailsGraphic = emailsFrame.Place(eps_file_path)
  # total emails number
  cirNum = myPage.TextFrames.Add()
  cirNum.GeometricBounds = graphic['cirNum']
  cirNum.ParentStory.Contents = '{:,}'.format(int(emails))
  cirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle
  # circle text
  cirText = myPage.TextFrames.Add()
  cirText.GeometricBounds = graphic['cirText']
  cirText.ParentStory.Contents = recent_mo + "\nTotal Email Recipients"
  cirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle
  # circle/product type heading
  cirHdg = myPage.TextFrames.Add()
  cirHdg.GeometricBounds = graphic['cirHdg']
  cirHdg.ParentStory.Contents = graphic['name']
  cirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle
  # data beneath it
  dataText = myPage.TextFrames.Add()
  dataText.GeometricBounds = graphic['dataText']
  dataText.ParentStory.Contents = str(open_rate) + "% Opens | " + str(open_change) + "% Change\n" + str(click_rate) + "% Clicks | " + str(click_change) + "% Change"
  dataText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle
  # arrow up or down
  arrow = myPage.GraphicLines.Add()
  arrow.GeometricBounds = graphic['arrow']
  arrow.StrokeColor = coDBlue
  arrow.StrokeWeight = 4
  arrow.LeftLineEnd = left_arrow_test(perc)
  arrow.RightLineEnd = right_arrow_test(perc)
  # percentage changed number
  percChanged = myPage.TextFrames.Add()
  percChanged.GeometricBounds = graphic['percChanged']
  percChanged.ParentStory.Contents = str(perc) + "%"
  percChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle
  # compared to text
  comText = myPage.TextFrames.Add()
  comText.GeometricBounds = graphic['comText']
  comText.ParentStory.Contents = "Compared to " + prev_mo + " Sent Emails"
  comText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle
  # compared to number
  comNum = myPage.TextFrames.Add()
  comNum.GeometricBounds = graphic['comNum']
  comNum.ParentStory.Contents = '{:,}'.format(int(prev_emails))
  comNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# Now let's do social media
# Social Media Subheading
socialRect = myPage.Rectangles.Add()
socialRect.GeometricBounds = geo_bounds_misc['socialRect']
socialRect.FillColor = coDBlue
socialRect.StrokeColor = strokeNone
socialSubH = myPage.TextFrames.Add()
socialSubH.GeometricBounds = geo_bounds_misc['socialSubH']
socialSubH.ParentStory.Contents = "Social Media"
socialSubH.ParentStory.Characters.Item(1).appliedParagraphStyle = subheadingStyle

# social impressions calcs
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
facebook_posts = facebook_posts.drop(facebook_posts.index[0]).reset_index(drop=True).rename(columns={'Post Message':'Post_Nickname','Lifetime Post Total Impressions':'Impressions','Lifetime Post Audience Targeting Unique Consumptions by Type - link clicks':'Clicks','Lifetime Matched Audience Targeting Consumptions on Post':'Interactions'})
facebook_posts['Posted'] = pd.to_datetime(facebook_posts['Posted'])
facebook_posts['Posted Month'] = facebook_posts['Posted'].dt.month
recent_facebook_posts = facebook_posts[facebook_posts['Posted Month'] == rec_month]
prev_facebook_posts = facebook_posts[facebook_posts['Posted Month'] == prev_month]
facebook_data_columns = ['Post_Nickname','Impressions','Clicks','Interactions']
recent_facebook_data = recent_facebook_posts[facebook_data_columns].assign(Source='Facebook').fillna(value=0).reset_index(drop=True)
prev_facebook_data = prev_facebook_posts[facebook_data_columns].assign(Source='Facebook').fillna(value=0).reset_index(drop=True)

recent_twitter_posts = pd.read_csv(rec_twitter_csv)
twitter_data_columns = ['Tweet text','impressions','url clicks','engagements']
recent_twitter_data = recent_twitter_posts[twitter_data_columns].assign(Source='Twitter').rename(columns={'Tweet text':'Post_Nickname','impressions':'Impressions','url clicks':'Clicks','engagements':'Interactions'})
prev_twitter_posts = pd.read_csv(prev_twitter_csv)
prev_twitter_data = prev_twitter_posts[twitter_data_columns].assign(Source='Twitter').rename(columns={'Tweet text':'Post_Nickname','impressions':'Impressions','url clicks':'Clicks','engagements':'Interactions'})

insta_posts = pd.read_csv(instagram_csv, engine='python')
insta_posts = insta_posts.assign(Clicks=0).assign(Interactions=(insta_posts.Likes + insta_posts.Comments)).assign(Source='Instagram').rename(columns={'Reach':'Impressions'})
insta_posts['Post_Date'] = pd.to_datetime(insta_posts['Post_Date'])
insta_posts['Posted Month'] = insta_posts['Post_Date'].dt.month
recent_insta_posts = insta_posts[insta_posts['Posted Month'] == rec_month]
prev_insta_posts = insta_posts[insta_posts['Posted Month'] == prev_month]
recent_insta_data = recent_insta_posts[['Post_Nickname','Impressions','Clicks','Interactions','Source']]

recent_posts_data = pd.concat([recent_facebook_data,recent_twitter_data,recent_insta_data],sort=True, ignore_index=True)
prev_posts_data = pd.concat([prev_facebook_data,prev_twitter_data,prev_insta_posts],sort=True, ignore_index=True)
recent_posts_data['Impressions'] = pd.to_numeric(recent_posts_data['Impressions'],downcast='integer').reset_index(drop=True)
recent_posts_data['Clicks'] = pd.to_numeric(recent_posts_data['Clicks'],downcast='integer').reset_index(drop=True)
recent_posts_data['Interactions'] = pd.to_numeric(recent_posts_data['Interactions'],downcast='integer').reset_index(drop=True)

recent_impressions_total = recent_posts_data['Impressions'].sum()
prev_posts_data['Impressions'] = pd.to_numeric(prev_posts_data['Impressions'],downcast='integer').reset_index(drop=True)
prev_impressions_total = prev_posts_data['Impressions'].sum()
impressions_perc_change = int(((recent_impressions_total-prev_impressions_total)/prev_impressions_total*100))

figSoc, axSoc = plt.subplots(figsize=(2.3,2.3), subplot_kw=dict(aspect="equal"))
axSoc.add_artist(outer_perc_change_wedge(0.5,impressions_perc_change))
axSoc.add_artist(outer_white_circle(0.435,impressions_perc_change))
axSoc.add_artist(perc_change_wedge((0.5,0.5),0.425,impressions_perc_change))
axSoc.add_artist(white_circle((0.5,0.5),0.33))
axSoc.add_artist(blue_circle((0.5,0.5),0.32))

plt.axis('off')
plt.tight_layout()
figSoc.savefig('social-impressions.eps',transparent=True)
socialImpEPS = cwd + '\\social-impressions.eps'

socialFrame = myPage.Rectangles.Add()
socialFrame.GeometricBounds = geo_bounds_misc['socialFrame']
socialFrame.StrokeColor = strokeNone
socialFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
socialGraphic = socialFrame.Place(socialImpEPS)

socialCirNum = myPage.TextFrames.Add()
socialCirNum.GeometricBounds = geo_bounds_misc['socialCirNum']
socialCirNum.ParentStory.Contents = '{:,}'.format(recent_impressions_total)
socialCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

socialCirText = myPage.TextFrames.Add()
socialCirText.GeometricBounds = geo_bounds_misc['socialCirText']
socialCirText.ParentStory.Contents = recent_mo + "\nImpressions"
socialCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

socialArrow = myPage.GraphicLines.Add()
socialArrow.GeometricBounds = geo_bounds_misc['socialArrow']
socialArrow.StrokeColor = coDBlue
socialArrow.StrokeWeight = 4
socialArrow.LeftLineEnd = left_arrow_test(impressions_perc_change)
socialArrow.RightLineEnd = right_arrow_test(impressions_perc_change)

socialPercChanged = myPage.TextFrames.Add()
socialPercChanged.GeometricBounds = geo_bounds_misc['socialPercChanged']
socialPercChanged.ParentStory.Contents = str(impressions_perc_change) + "%"
socialPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

socialComText = myPage.TextFrames.Add()
socialComText.GeometricBounds = geo_bounds_misc['socialComText']
socialComText.ParentStory.Contents = "Compared to " + prev_mo + " Total"
socialComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

socialComNum = myPage.TextFrames.Add()
socialComNum.GeometricBounds = geo_bounds_misc['socialComNum']
socialComNum.ParentStory.Contents = '{:,}'.format(prev_impressions_total)
socialComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# now for followers and posts counts
count_twitter_followers = int(inputs_df.iloc[5]['CSV File Name or Numbers'])
count_twitter_posts = recent_twitter_data['Post_Nickname'].count()
count_fb_followers = int(inputs_df.iloc[6]['CSV File Name or Numbers'])
count_fb_posts = recent_facebook_data['Post_Nickname'].count()
count_insta_followers = int(inputs_df.iloc[7]['CSV File Name or Numbers'])
count_insta_posts = recent_insta_data['Post_Nickname'].count()

twitterIconFrame = myPage.Rectangles.Add()
twitterIconFrame.GeometricBounds = geo_bounds_misc['twitterIconFrame']
twitterIconFrame.StrokeColor = strokeNone
twitterIconFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
twitterIcon = twitterIconFrame.Place(twitterEPS)
twitterData = myPage.TextFrames.Add()
twitterData.GeometricBounds = geo_bounds_misc['twitterData']
twitterData.ParentStory.Contents = "Followers | "+ '{:,}'.format(count_twitter_followers) + "\nPosts | " + '{:,}'.format(count_twitter_posts)
twitterData.ParentStory.Characters.Item(1).appliedParagraphStyle = socialDataStyle

facebookIconFrame = myPage.Rectangles.Add()
facebookIconFrame.GeometricBounds = geo_bounds_misc['facebookIconFrame']
facebookIconFrame.StrokeColor = strokeNone
facebookIconFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
facebookIcon = facebookIconFrame.Place(facebookEPS)
facebookData = myPage.TextFrames.Add()
facebookData.GeometricBounds = geo_bounds_misc['facebookData']
facebookData.ParentStory.Contents = "Followers | "+ '{:,}'.format(count_fb_followers) + "\nPosts | " + '{:,}'.format(count_fb_posts)
facebookData.ParentStory.Characters.Item(1).appliedParagraphStyle = socialDataStyle

instagramIconFrame = myPage.Rectangles.Add()
instagramIconFrame.GeometricBounds = geo_bounds_misc['instagramIconFrame']
instagramIconFrame.StrokeColor = strokeNone
instagramIconFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
instagramIcon = instagramIconFrame.Place(instagramEPS)
instagramData = myPage.TextFrames.Add()
instagramData.GeometricBounds =geo_bounds_misc['instagramData']
instagramData.ParentStory.Contents = "Followers | "+ '{:,}'.format(count_insta_followers) + "\nPosts | " + '{:,}'.format(count_insta_posts)
instagramData.ParentStory.Characters.Item(1).appliedParagraphStyle = socialDataStyle

# top performing social posts
recent_posts_data = recent_posts_data.sort_values(by=['Interactions'],ascending=False).reset_index(drop=True)
top_social_posts = recent_posts_data.iloc[:5].rename(columns={'Post_Nickname':'Social Post','Source':'Network'})
top_social_posts = top_social_posts[['Social Post','Network','Impressions','Clicks','Interactions']].replace(regex=True, to_replace="'", value="").replace(regex=True, to_replace="\n", value="")
top_social_posts['Impressions'] = top_social_posts['Impressions'].apply(lambda x: '{:,}'.format(x))
top_social_posts['Clicks'] = top_social_posts['Clicks'].apply(lambda x: '{:,}'.format(x))
top_social_posts['Interactions'] = top_social_posts['Interactions'].apply(lambda x: '{:,}'.format(x))
top_social_posts_list = df_to_list(top_social_posts)

socialTableText = myPage.TextFrames.Add()
socialTableText.GeometricBounds = geo_bounds_misc['socialTableText']
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
socialTable.Columns.Item(4).FillColor = coYellow
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

# Now for website visits!
# Website Visits Subheading
webRect = myPage.Rectangles.Add()
webRect.GeometricBounds = geo_bounds_misc['webRect']
webRect.FillColor = coDBlue
webRect.StrokeColor = strokeNone
webSubH = myPage.TextFrames.Add()
webSubH.GeometricBounds = geo_bounds_misc['webSubH']
webSubH.ParentStory.Contents = "Website Visits"
webSubH.ParentStory.Characters.Item(1).appliedParagraphStyle = subheadingStyle

# visits by source calcs
charleston_traffic_csv = inputs_df.iloc[8]['CSV File Name or Numbers']
while '.csv' not in charleston_traffic_csv:
    charleston_traffic_csv = input('Please provide the name of the CSV file, including the .csv, for the month of Charleston website data: ')
columbia_traffic_csv = inputs_df.iloc[9]['CSV File Name or Numbers']
while '.csv' not in columbia_traffic_csv:
    columbia_traffic_csv = input('Please provide the name of the CSV file, including the .csv, for the month of Columbia website data: ')
greenville_traffic_csv = inputs_df.iloc[10]['CSV File Name or Numbers']
while '.csv' not in greenville_traffic_csv:
    greenville_traffic_csv = input('Please provide the name of the CSV file, including the .csv, for the month of Greenville website data: ')
hs_traffic_csv = inputs_df.iloc[11]['CSV File Name or Numbers']
while '.csv' not in hs_traffic_csv:
    hs_traffic_csv = input('Please provide the name of the CSV file, including the .csv, for the month of HubSpot website data: ')
    
chs_traffic_df = pd.read_csv(charleston_traffic_csv,skiprows=5)
stop_chs = np.where(chs_traffic_df['Page'].isna())
chs_traffic_df = chs_traffic_df.iloc[:(stop_chs[0])[0]]
cae_traffic_df = pd.read_csv(columbia_traffic_csv,skiprows=5)
stop_cae = np.where(cae_traffic_df['Page'].isna())
cae_traffic_df = cae_traffic_df.iloc[:(stop_cae[0])[0]]
grv_traffic_df = pd.read_csv(greenville_traffic_csv,skiprows=5)
stop_grv = np.where(grv_traffic_df['Page'].isna())
grv_traffic_df = grv_traffic_df.iloc[:(stop_grv[0])[0]]
co_traffic_df = pd.concat([chs_traffic_df,cae_traffic_df,grv_traffic_df],sort=True, ignore_index=True)

hs_page_traffic = pd.read_csv(hs_traffic_csv,skiprows=5)
stop_h = np.where(hs_page_traffic['Page'].isna())
hs_page_traffic = hs_page_traffic.iloc[:(stop_h[0])[0]]

co_page_views = co_traffic_df[['Page','Source / Medium','Pageviews']]
co_page_views.Pageviews = pd.to_numeric(co_page_views.Pageviews,downcast='integer')
hs_page_views = hs_page_traffic[['Page','Source / Medium','Pageviews']]
hs_page_views.Pageviews = pd.to_numeric(hs_page_views.Pageviews,downcast='integer')
co_page_views_totals = co_page_views.Pageviews.sum() + hs_page_views.Pageviews.sum()
total_page_views = '{:,}'.format(co_page_views_totals)

co_web_sources = pd.concat([co_page_views,hs_page_views],sort=True, ignore_index=True)
co_web_sources['Source / Medium'] = co_web_sources['Source / Medium'].replace('linkedin.com / referral','linkedin.com / social post').replace('m.facebook.com / referral','m.facebook.com / social post').replace('lnkd.in / referral','lnkd.in / social post').replace('facebook.com / referral','facebook.com / social post')
co_web_sources = co_web_sources.assign(Source=co_web_sources['Source / Medium'].apply(lambda s: s.split('/ ')[1]))
co_web_sources['Source'] = co_web_sources['Source'].replace('(none)','direct').replace('(not set)','referral').replace('social post','social')
co_web_sources_group = co_web_sources.groupby('Source').Pageviews.sum().reset_index()
co_web_sources_group.Source = co_web_sources_group.Source.str.capitalize()
co_web_sources_group['Percentage'] = co_web_sources_group.Pageviews.apply(lambda x: str(int(100*x/co_page_views_totals))).reset_index(drop=True)
co_web_sources_group = co_web_sources_group[co_web_sources_group.Percentage != '0']
co_web_sources_group['Label'] = co_web_sources_group.apply(lambda row: row.Percentage + "%\n" + row.Source, axis=1)
sessions_count = co_web_sources_group['Pageviews'].tolist()
channel_labels = co_web_sources_group['Label'].tolist()
figSrc, axSrc = plt.subplots(figsize=(3.1,3.1), subplot_kw=dict(aspect="equal"))
wedgesSrc,textsSrc = axSrc.pie(co_web_sources_group['Pageviews'].tolist(),
                        colors=co_colors,
                        labels=co_web_sources_group['Label'].tolist(),
                        labeldistance=1.2,
                        startangle=90,
                        wedgeprops=dict(width=0.25,linewidth=2,edgecolor='w'))
for t in textsSrc:
    t.set_horizontalalignment('center')
axSrc.add_artist(blue_circle((0,0),0.7))
plt.tight_layout()
figSrc.savefig('website-visits.eps',transparent=True)
websiteVisitsEPS = cwd + '\\website-visits.eps'

websiteFrame = myPage.Rectangles.Add()
websiteFrame.GeometricBounds = geo_bounds_misc['websiteFrame']
websiteFrame.StrokeColor = strokeNone
websiteFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
websiteGraphic = websiteFrame.Place(websiteVisitsEPS)

websiteCirNum = myPage.TextFrames.Add()
websiteCirNum.GeometricBounds = geo_bounds_misc['websiteCirNum']
websiteCirNum.ParentStory.Contents = total_page_views
websiteCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = largeCircleNumbersStyle

websiteCirText = myPage.TextFrames.Add()
websiteCirText.GeometricBounds = geo_bounds_misc['websiteCirText']
websiteCirText.ParentStory.Contents = recent_mo + "\nTotal Website Visits"
websiteCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

# website visits by campaign
sessions_by_campaigns_csv = inputs_df.iloc[12]['CSV File Name or Numbers']
while '.csv' not in sessions_by_campaigns_csv:
    sessions_by_campaigns_csv = input('Please provide the name of the CSV file, including the .csv, for the month of sessions data: ')

campaign_sessions_df = pd.read_csv(sessions_by_campaigns_csv)
campaign_sessions_df = campaign_sessions_df.drop(campaign_sessions_df.index[-2:]).sort_values(by=['Name'],ascending=True).reset_index(drop=True)
campaign_sessions_df = campaign_sessions_df[(campaign_sessions_df.Name.isnull() == False) & (campaign_sessions_df.Name.str.contains('Offline') == False)]
campaign_sessions_df['Total Visits'] = campaign_sessions_df.apply(lambda row: int(sum(row[1:])),axis=1)
campaign_visits = campaign_sessions_df[['Name','Total Visits']].reset_index(drop=True)
campaign_visits = campaign_visits.replace(regex=True, to_replace="Columbia |Greenville |Charleston", value="").reset_index(drop=True)
campaign_visits['Name'] = campaign_visits.Name.replace(regex=True, to_replace="\d+", value="").apply(lambda n: n.split(' - ')[0])
campaign_visits_group = campaign_visits.groupby('Name')['Total Visits'].sum().reset_index()
campaign_visits_group = campaign_visits_group[(campaign_visits_group.Name != 'REMS Property Newsletters') & (campaign_visits_group.Name.str.contains('Internal') == False) & (campaign_visits_group.Name.str.contains('Sample') == False) & (campaign_visits_group['Total Visits'] != 0)].sort_values(by=['Total Visits'],ascending=False).reset_index(drop=True)
campaign_visits_group['Total Visits'] = campaign_visits_group['Total Visits'].apply(lambda x: '{:,}'.format(x))
campaign_visits_list = df_to_list(campaign_visits_group)

campaignTableText = myPage.TextFrames.Add()
campaignTableText.GeometricBounds = geo_bounds_misc['campaignTableText']
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
campaignTable.Columns.Item(2).FillColor = coYellow
campaignTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
campaignTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
campaignTable.Rows.Item(1).VerticalJustification = myVAlignBottom
campaignTable.Contents = campaign_visits_list
campaignrangetotal = campaigncolumncount * campaignrowcount
for i in range(2,campaignrowcount):
    campaignTable.Rows.Item(i).VerticalJustification = myVAlignCenter
for i in range(1,campaignrangetotal,campaigncolumncount):
    campaignTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign

# visits by page type
co_page_types = co_page_views[['Page','Pageviews']]
co_page_types.Page = co_page_types.Page.str.lower()
co_page_types = co_page_types.replace(regex=True, to_replace="/en", value="").replace(regex=True, to_replace="/united-states", value="").reset_index(drop=True)
co_page_types['Page Type'] = co_page_types['Page'].apply(lambda p: p.split('/')[1]).apply(lambda p: p.capitalize())

hs_page_types = hs_page_views[['Page','Pageviews']]
hs_page_types.Page = hs_page_types.Page.str.lower()
hs_page_types['Page Type'] = hs_page_types['Page'].apply(hs_page_type)

all_page_types = pd.concat([co_page_types,hs_page_types],sort=True, ignore_index=True)
co_page_groups = all_page_types.groupby('Page Type').Pageviews.sum().reset_index()
co_page_groups = co_page_groups[(co_page_groups['Page Type'].str.contains('cache') == False) & (co_page_groups['Page Type'].str.contains('_hcms') == False) & (co_page_groups['Page Type'].str.contains('translate') == False) & (co_page_groups['Page Type'].str.contains('edit') == False)].sort_values(by=['Pageviews'],ascending=False).reset_index(drop=True)
co_page_groups.Pageviews = co_page_groups.Pageviews.apply(lambda x: '{:,}'.format(x))
co_page_groups_list = df_to_list(co_page_groups)

pagesTableText = myPage.TextFrames.Add()
pagesTableText.GeometricBounds = geo_bounds_misc['pagesTableText']
pagesTableText.ParentStory.Contents = "Which pages were they visiting?"
pagesTableText.ParentStory.Characters.Item(1).appliedParagraphStyle = tableTitleStyle
pagesTable = pagesTableText.ParentStory.InsertionPoints.Item(-1).Tables.Add()
pagescolumncount = 2
pagesrowcount = len(co_page_groups) + 1
pagesTable.ColumnCount = pagescolumncount
pagesTable.HeaderRowCount = 1
pagesTable.BodyRowCount = pagesrowcount - 1
pagesTable.Height = '2i'
pagesTable.appliedTableStyle = myTableStyle
pagesTable.Columns.Item(1).Width = '1.5356i'
pagesTable.Columns.Item(2).Width = '0.8142i'
pagesTable.Columns.Item(2).FillColor = coYellow
pagesTable.Rows.Item(-1).BottomEdgeStrokeColor = strokeNone
pagesTable.Rows.Item(-1).BottomEdgeStrokeWeight = 0
pagesTable.Rows.Item(1).VerticalJustification = myVAlignBottom
for i in range(2,7):
    pagesTable.Rows.Item(i).VerticalJustification = myVAlignCenter
pagesTable.Contents = co_page_groups_list
pagesrangetotal = pagescolumncount * pagesrowcount
for i in range(1,pagesrangetotal,pagescolumncount):
    pagesTable.Cells.Item(i).Texts.Item(1).Justification = myLeftAlign

# Property subscribers and research subscribers
current_prop_subscribers = '{:,}'.format(int(inputs_df.iloc[12]['CSV File Name or Numbers']))
current_research_subscribers = '{:,}'.format(int(inputs_df.iloc[13]['CSV File Name or Numbers']))

propSubText = myPage.TextFrames.Add()
propSubText.GeometricBounds = geo_bounds_misc['propSubText']
propSubText.ParentStory.Contents = current_prop_subscribers
propSubText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
propSubHdg = myPage.TextFrames.Add()
propSubHdg.GeometricBounds = geo_bounds_misc['propSubHdg']
propSubHdg.ParentStory.Contents = "Property Subscriber Leads"
propSubHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

resSubText = myPage.TextFrames.Add()
resSubText.GeometricBounds = geo_bounds_misc['resSubText']
resSubText.ParentStory.Contents = current_research_subscribers
resSubText.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataStyle
resSubHdg = myPage.TextFrames.Add()
resSubHdg.GeometricBounds = geo_bounds_misc['resSubHdg']
resSubHdg.ParentStory.Contents = "Research Subscriber Leads"
resSubHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = plainDataSubheadingStyle

# Marketing Initiatives
mktgRect = myPage.Rectangles.Add()
mktgRect.GeometricBounds = geo_bounds_misc['mktgRect']
mktgRect.FillColor = coDBlue
mktgRect.StrokeColor = strokeNone
mktgSubH = myPage.TextFrames.Add()
mktgSubH.GeometricBounds = geo_bounds_misc['mktgSubH']
mktgSubH.ParentStory.Contents = "Marketing Initiatives"
mktgSubH.ParentStory.Characters.Item(1).appliedParagraphStyle = subheadingStyle

# we're going to do these as a loop as well, so let's start with preliminary calculations
ytd_prs = int(inputs_df.iloc[15]['CSV File Name or Numbers'])
prev_ytd_prs = int(inputs_df.iloc[16]['CSV File Name or Numbers'])

qtrly_emails_csv = inputs_df.iloc[0]['CSV File Name or Numbers']
while '.csv' not in qtrly_emails_csv:
    qtrly_emails_csv = input('Please provide the name of the CSV file, including the .csv, for the email data being compared: ')
mr_emails_df = pd.read_csv(qtrly_emails_csv)
mr_emails_df = mr_emails_df[mr_emails_df['Campaign'].str.contains('Market Reports')==True].reset_index(drop=True)
mr_emails_df = mr_emails_df[mr_emails_df['Email Name'].str.startswith('20')==True].reset_index(drop=True)
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

current_loopnet_views = int(inputs_df.iloc[17]['CSV File Name or Numbers'])
prev_loopnet_views = int(inputs_df.iloc[18]['CSV File Name or Numbers'])

current_proposals_csv = inputs_df.iloc[19]['CSV File Name or Numbers']
while '.csv' not in current_proposals_csv:
    current_proposals_csv = input('Please provide the name of the CSV file, including the .csv, for this year\'s tracking report: ')
prev_proposals_csv = inputs_df.iloc[20]['CSV File Name or Numbers']
while '.csv' not in prev_proposals_csv:
    prev_proposals_csv = input('Please provide the name of the CSV file, including the .csv, for last year\'s tracking report: ')
current_proposals_df = pd.read_csv(current_proposals_csv)
current_proposals_df = current_proposals_df.fillna(value=1).reset_index(drop=True)
current_proposals_df['Submission_Month'] = pd.to_numeric(current_proposals_df['Submission Date'].apply(lambda x: str(x).split('.')[0]))
current_month = current_proposals_df['Submission_Month'].max()
prev_proposals_df = pd.read_csv(prev_proposals_csv)
prev_proposals_df = prev_proposals_df.fillna(value=1).reset_index(drop=True)
prev_proposals_df['Submission_Month'] = pd.to_numeric(prev_proposals_df['Submission Date'].apply(lambda x: str(x).split('.')[0]))
prev_proposals_ytd = prev_proposals_df[prev_proposals_df.Submission_Month <= current_month]
ytd_proposals = current_proposals_df['Submission_Month'].count()
prev_ytd_proposals = prev_proposals_ytd['Submission_Month'].count()

# then our dictionary
init_graphics = {}
init_graphics['press'] = {}
init_graphics['press']['name'] = 'Press Releases'
init_graphics['press']['cirHdg'] = ["8.2337i","0.2079i","8.375i","1.9779i"]
init_graphics['press']['eps_file'] = 'press-releases.eps'
init_graphics['press']['frame'] = ["6.2222i","-0.1439i","8.5139i","2.1478i"]
init_graphics['press']['circle_data'] = (str(ytd_prs))
init_graphics['press']['cirNum'] = ["7.0571i","0.6915i","7.3025i","1.4943i"]
init_graphics['press']['circle_text'] = 'YTD Press Releases'
init_graphics['press']['cirText'] = ["7.315i","0.6915i","7.5659i","1.4943i"]
init_graphics['press']['change'] = (int(((ytd_prs-prev_ytd_prs)/prev_ytd_prs*100)))
init_graphics['press']['arrow'] = ["6.6423i","2.3656i","6.9001i","2.3656i"]
init_graphics['press']['percChanged'] = ["6.9965i","2.0505i","7.2399i","2.6806i"]
init_graphics['press']['compared_text'] = 'YTD 2017'
init_graphics['press']['comText'] = ["7.2399i","2.0505i","7.7138i","2.6806i"]
init_graphics['press']['compared_data'] = (str(prev_ytd_prs))
init_graphics['press']['comNum'] = ["7.7138i","2.0505i","7.857i","2.6806i"]
init_graphics['reports'] = {}
init_graphics['reports']['name'] = 'Market Reports'
init_graphics['reports']['cirHdg'] = ["8.2337i","2.8746i","8.375i","4.6446i"]
init_graphics['reports']['eps_file'] = 'market-reports.eps'
init_graphics['reports']['frame'] = ["6.2222i","2.5255i","8.5139i","4.8171i"]
init_graphics['reports']['circle_data'] = ('{:,}'.format(recent_mr_recipients))
init_graphics['reports']['cirNum'] = ["6.9965i","3.3582i","7.2419i","4.161i"]
init_graphics['reports']['circle_text'] = recent_qtr + ' Total Email Recipients'
init_graphics['reports']['cirText'] = ["7.2408i","3.3582i","7.64i","4.161i"]
init_graphics['reports']['change'] = (int(((recent_mr_recipients - prev_mr_recipients)/prev_mr_recipients)*100))
init_graphics['reports']['arrow'] = ["6.6423i","5.0322i","6.9001i","5.0322i"]
init_graphics['reports']['percChanged'] = ["6.9965i","4.7172i","7.2399i","5.3472i"]
init_graphics['reports']['compared_text'] = '\n' + prev_qtr
init_graphics['reports']['comText'] = ["7.2399i","4.7172i","7.7138i","5.3472i"]
init_graphics['reports']['compared_data'] = '{:,}'.format(prev_mr_recipients)
init_graphics['reports']['comNum'] = ["7.7138i","4.7172i","7.857i","5.3472i"]

# now the loop
for init_graphic in init_graphics.values():
  fig,ax = plt.subplots(figsize=(2.3,2.3), subplot_kw=dict(aspect="equal"))
  ax.add_artist(outer_perc_change_wedge(0.5,init_graphic['change']))
  ax.add_artist(outer_white_circle(0.435,init_graphic['change']))
  ax.add_artist(perc_change_wedge((.5,.5),0.425,init_graphic['change']))
  ax.add_artist(white_circle((.5,.5),0.33))
  ax.add_artist(blue_circle((.5,.5),0.32))
  plt.axis('off')
  plt.tight_layout()
  fig.savefig(init_graphic['eps_file'],transparent=True)
  eps_file_path = cwd + '\\' + init_graphic['eps_file']
  # now for placing in the file
  frame = myPage.Rectangles.Add()
  frame.GeometricBounds = init_graphic['frame']
  frame.StrokeColor = strokeNone
  frame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
  graphic = frame.Place(eps_file_path)
  # circle number
  cirNum = myPage.TextFrames.Add()
  cirNum.GeometricBounds = init_graphic['cirNum']
  cirNum.ParentStory.Contents = init_graphic['circle_data']
  cirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle
  # circle text
  cirText = myPage.TextFrames.Add()
  cirText.GeometricBounds = init_graphic['cirText']
  cirText.ParentStory.Contents = init_graphic['circle_text']
  cirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle
  # circle/data type heading
  cirHdg = myPage.TextFrames.Add()
  cirHdg.GeometricBounds = init_graphic['cirHdg']
  cirHdg.ParentStory.Contents = init_graphic['name']
  cirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle
  # arrow up or down
  arrow = myPage.GraphicLines.Add()
  arrow.GeometricBounds = init_graphic['arrow']
  arrow.StrokeColor = coDBlue
  arrow.StrokeWeight = 4
  arrow.LeftLineEnd = left_arrow_test(init_graphic['change'])
  arrow.RightLineEnd = right_arrow_test(init_graphic['change'])
  # percent changed
  percChanged = myPage.TextFrames.Add()
  percChanged.GeometricBounds = init_graphic['percChanged']
  percChanged.ParentStory.Contents = str(init_graphic['change'])
  percChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle
  # compared to text
  comText = myPage.TextFrames.Add()
  comText.GeometricBounds = init_graphic['comText']
  comText.ParentStory.Contents = 'Compared to ' + init_graphic['compared_text']
  comText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle
  # compared to number
  comNum = myPage.TextFrames.Add()
  comNum.GeometricBounds = init_graphic['comNum']
  comNum.ParentStory.Contents = init_graphic['compared_data']
  comNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

  # now for advertising
recent_facebook_posts['Lifetime Post Paid Impressions'] = pd.to_numeric(recent_facebook_posts['Lifetime Post Paid Impressions'],downcast='integer').reset_index(drop=True)
recent_facebook_posts['Interactions'] = pd.to_numeric(recent_facebook_posts['Interactions'],downcast='integer').reset_index(drop=True)
recent_facebook_ads = recent_facebook_posts[recent_facebook_posts['Lifetime Post Paid Impressions']>0].reset_index(drop=True)
facebook_ad_columns = ['Post ID','Lifetime Post Paid Impressions','Interactions']
recent_facebook_ads = recent_facebook_ads[facebook_ad_columns].rename(columns={'Lifetime Post Paid Impressions':'Impressions'}).reset_index(drop=True)
total_ad_impressions = recent_facebook_ads.Impressions.sum()
total_ad_interactions = recent_facebook_ads.Interactions.sum()
try:
    ad_interactions_perc = int((total_ad_interactions/total_ad_impressions)*100)
except ValueError:
    ad_interactions_perc = 0
figAd,axAd = plt.subplots(figsize=(2.3,2.3), subplot_kw=dict(aspect="equal"))
axAd.add_artist(perc_change_wedge((.5,.5),0.425,ad_interactions_perc))
axAd.add_artist(white_circle((.5,.5),0.33))
axAd.add_artist(blue_circle((.5,.5),0.32))
plt.axis('off')
plt.tight_layout()
figAd.savefig('ad-views.eps',transparent=True)
adEPS = cwd + '\\ad-views.eps'

adFrame = myPage.Rectangles.Add()
adFrame.GeometricBounds = geo_bounds_misc['adFrame']
adFrame.StrokeColor = strokeNone
adFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
adGraphic = adFrame.Place(adEPS)

adCirNum = myPage.TextFrames.Add()
adCirNum.GeometricBounds = geo_bounds_misc['adCirNum']
adCirNum.ParentStory.Contents = '{:,}'.format(total_ad_impressions)
adCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

adCirText = myPage.TextFrames.Add()
adCirText.GeometricBounds = geo_bounds_misc['adCirText']
adCirText.ParentStory.Contents = 'Total Ad Impressions'
adCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

adCirHdg = myPage.TextFrames.Add()
adCirHdg.GeometricBounds = geo_bounds_misc['adCirHdg']
adCirHdg.ParentStory.Contents = 'Advertising'
adCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

adPercChanged = myPage.TextFrames.Add()
adPercChanged.GeometricBounds = geo_bounds_misc['adPercChanged']
adPercChanged.ParentStory.Contents = str(ad_interactions_perc) + '%'
adPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

adComText = myPage.TextFrames.Add()
adComText.GeometricBounds = geo_bounds_misc['adComText']
adComText.ParentStory.Contents = 'Interactions from Impressions'
adComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

adComNum = myPage.TextFrames.Add()
adComNum.GeometricBounds = geo_bounds_misc['adComNum']
adComNum.ParentStory.Contents = '{:,}'.format(total_ad_interactions) + ' Total'
adComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle

# finally, proposals
current_proposals_csv = inputs_df.iloc[17]['CSV File Name or Numbers']
while '.csv' not in current_proposals_csv:
    current_proposals_csv = input('Please provide the name of the CSV file, including the .csv, for this year\'s tracking report: ')
current_proposals_df = pd.read_csv(current_proposals_csv)
current_proposals_df = current_proposals_df.assign(Outcome=current_proposals_df.apply(proposal_outcome,axis=1))
grouped_proposals = current_proposals_df.groupby('Outcome')['Property/Proposal Name'].count().reset_index().rename(columns={'Property/Proposal Name':'Proposals'})
total_proposals = grouped_proposals.Proposals.sum()
grouped_proposals['Percentage'] = grouped_proposals['Proposals'].apply(lambda x: str(int(100*x/total_proposals))).reset_index(drop=True)
grouped_proposals['Label'] = grouped_proposals.apply(lambda row: row.Percentage + "%\n" + row.Outcome, axis=1)
proposals_count = grouped_proposals['Proposals'].tolist()
outcome_labels = grouped_proposals['Label'].tolist()
figProp, axProp = plt.subplots(figsize=(2.2,2.2), subplot_kw=dict(aspect="equal"))
wedgesProp,textsProp = axProp.pie(grouped_proposals['Proposals'].tolist(),
                        colors=colliers_colors,
                        labels=grouped_proposals['Label'].tolist(),
                        labeldistance=1.23,
                        startangle=90,
                        wedgeprops=dict(width=0.25,linewidth=2,edgecolor='w'))
for t in textsProp:
    t.set_horizontalalignment('center')
axProp.add_artist(blue_circle((0,0),0.7))
plt.tight_layout()
figProp.savefig('proposals.eps',transparent=True)
proposalsEPS = cwd + '\\proposals.eps'

lost_proposals = outcomes_output('Lost')
outstanding_proposals = outcomes_output('Outstanding')
#pulled_proposals = outcomes_output('Pulled')
won_proposals = outcomes_output('Won')

propFrame = myPage.Rectangles.Add()
propFrame.GeometricBounds = geo_bounds_misc['propFrame']
propFrame.StrokeColor = strokeNone
propFrame.FrameFittingOptions.FittingOnEmptyFrame = myEmptyFitProp
propGraphic = propFrame.Place(proposalsEPS)

propCirNum = myPage.TextFrames.Add()
propCirNum.GeometricBounds = geo_bounds_misc['propCirNum']
propCirNum.ParentStory.Contents = str(won_proposals)
propCirNum.ParentStory.Characters.Item(1).appliedParagraphStyle = circleNumbersStyle

propCirText = myPage.TextFrames.Add()
propCirText.GeometricBounds = geo_bounds_misc['propCirText']
propCirText.ParentStory.Contents = 'Proposals Won YTD'
propCirText.ParentStory.Characters.Item(1).appliedParagraphStyle = circleTextStyle

propCirHdg = myPage.TextFrames.Add()
propCirHdg.GeometricBounds = geo_bounds_misc['propCirHdg']
propCirHdg.ParentStory.Contents = 'Proposals'
propCirHdg.ParentStory.Characters.Item(1).appliedParagraphStyle = circleHeadingStyle

propPercChanged = myPage.TextFrames.Add()
propPercChanged.GeometricBounds = geo_bounds_misc['propPercChanged']
propPercChanged.ParentStory.Contents = 'YTD'
propPercChanged.ParentStory.Characters.Item(1).appliedParagraphStyle = percentChangedStyle

propComText = myPage.TextFrames.Add()
propComText.GeometricBounds = geo_bounds_misc['propComText']
propComText.ParentStory.Contents = str(lost_proposals) + ' Lost\n' + str(outstanding_proposals) + ' Out-\nstanding'
propComText.ParentStory.Characters.Item(1).appliedParagraphStyle = dataTextStyle

propComNum = myPage.TextFrames.Add()
propComNum.GeometricBounds = geo_bounds_misc['propComNum']
propComNum.ParentStory.Contents = '{:,}'.format(total_proposals) + ' Total'
propComNum.ParentStory.Characters.Item(1).appliedParagraphStyle = comparedNumberStyle
