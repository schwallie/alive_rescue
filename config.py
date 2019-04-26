from string import ascii_uppercase

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from secret import *


def df_to_google_doc(df, workbook_name, wks_name, include_col_names=True, include_index=True):
    """

    :param df: DataFrame to write
    :param workbook_name: Name of workbook
    :param wks_name: Name or worksheet
    :param include_col_names:
    :param include_index:
    :return:
    """
    wkb = open_connection_to_google_spreadsheet(workbook_name)
    wks = wkb.worksheet(wks_name)
    last_row = len(df.index) if include_col_names else len(df.index) - 1
    # Give final col name
    result = []
    col = len(df.columns)
    while col:
        col, rem = divmod(col - 1, 26)
        result[:0] = ascii_uppercase[rem]
    last_col = ''.join(result)
    # Give first col/row name
    first_col = 'B1' if include_index else 'A1'
    first_row = 'A2' if include_col_names else 'A1'

    # Add rows/cols if necessary
    if last_row + 1 > wks.row_count:
        wks.add_rows(last_row - wks.row_count + 2)

    if len(df.columns) + 1 > wks.col_count:
        wks.add_cols(len(df.columns) - wks.col_count + 2)

    # Add col names to sheet
    if include_col_names:
        cell_list = wks.range('%s:%s1' % (first_col, last_col))
        for idx, cell in enumerate(cell_list):
            cell.value = df.columns.values[idx]
        wks.update_cells(cell_list)

    # Add index to sheet if needed
    if include_index:
        cell_list = wks.range('%s:A%d' % (first_row, last_row + 1))
        for idx, cell in enumerate(cell_list):
            cell.value = df.index[idx]
        wks.update_cells(cell_list)

    # Add cell values to sheet
    # TODO: Large dataframes will break this
    cells = wks.range('%s%s:%s%d' % (first_col[0], first_row[1], last_col, last_row + 1))
    for ix_idx, idx in enumerate(df.index):
        for ix_col, col in enumerate(df.columns.values):
            col_multiplier = ix_idx * len(df.columns)
            cells[ix_col + col_multiplier].value = str(df[col][idx])
    wks.update_cells(cells)


def open_connection_to_google_spreadsheet(spreadsheet_name):
    """
    Opens a Google spreadsheet
    :param spreadsheet_name: The name of the spreadsheet
    :return: workbook object
    """
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(drive_details, scope)
    gc = gspread.authorize(credentials)
    return gc.open(spreadsheet_name)


def get_first_bday_of_month(mnth=None, yr=None):
    '''
    Return the first business day of the current month if no variables provided
    Return the first business day of the month and year provided if variables provided

    Tests:
    In [188]: config.get_first_bday_of_month(12,2015)
    Out[188]: datetime.date(2015, 12, 1)

    In [189]: config.get_first_bday_of_month(11,2015)
    Out[189]: datetime.date(2015, 11, 2)

    In [190]: config.get_first_bday_of_month(10,2015)
    Out[190]: datetime.date(2015, 10, 1)

    In [191]: config.get_first_bday_of_month(1,2016)
    Out[191]: datetime.date(2016, 1, 4)

    In [192]: config.get_first_bday_of_month(8,2015)
    Out[192]: datetime.date(2015, 8, 3)
    :param mnth:
    :param yr:
    :return:
    '''
    from calendar import monthrange
    from pandas.tseries.holiday import USFederalHolidayCalendar
    from pandas.tseries.offsets import CustomBusinessDay
    if yr is None or mnth is None:
        yr = pd.datetime.now().year if pd.datetime.now().month != 1 else pd.datetime.now().year - 1
        mnth = pd.datetime.now().month - 1 if pd.datetime.now().month != 1 else 12
    else:
        yr = yr if mnth != 1 else yr - 1
        mnth = mnth - 1 if mnth != 1 else 12
    end_last = monthrange(yr, mnth)
    end_last = pd.Timestamp('%s/%s/%s' % (mnth, end_last[1], yr)).date()
    cal = USFederalHolidayCalendar()
    holidays = cal.holidays(start=end_last - pd.tseries.offsets.Day(60),
                            end=end_last + pd.tseries.offsets.Day(60)).to_pydatetime()
    bday_cus = CustomBusinessDay(holidays=holidays)
    return (end_last + bday_cus).date()


def get_last_weekday_value(weekday):  # Mon = 0, Sun = 6
    '''
    If you want to get the last Friday, use weekday=4, last Thurs = 3, Sun = 6
    '''
    day_minus = pd.datetime.today().weekday() + weekday - 1
    return (pd.datetime.today() - pd.tseries.offsets.Day(day_minus)).date()


'''
##############
#
#Getting files and putting files to dropbox.
#
##############
'''


def format_percent(item):
    return '{:.1%}'.format(item)


def format_money(item, locale='us'):
    if locale == 'uk':
        return u'\xA3{:,.0f}'.format(item)
    else:
        return '${:,.0f}'.format(item)


def format_money_gbp(item, locale='uk'):
    if locale == 'uk':
        return u'\xA3{:,.0f}'.format(item)
    else:
        return '${:,.0f}'.format(item)


def format_money_w_decimal(item, locale='us'):
    if locale == 'uk':
        return u'\xA3{:,.0f}'.format(item)
    else:
        return '${:,.2f}'.format(item)


def format_number(item):
    return '{:,.0f}'.format(item)


def format_number_w_decimal(item):
    return '{:,.1f}'.format(item)


__author__ = "Chase Schwalbach"
__credits__ = ["Chase Schwalbach"]
__version__ = "1.0"
__maintainer__ = "Chase Schwalbach"
__email__ = "chase.schwalbach@avant.com"
__status__ = "Production"
