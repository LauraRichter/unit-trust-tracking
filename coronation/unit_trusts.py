
import requests
import time
import os
import pandas as pd
import numpy as np
import matplotlib as mpl
mpl.use('Agg')
import matplotlib.pylab as plt

import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
import httplib2
import oauth2client
from oauth2client import client, tools
import base64
from apiclient import errors, discovery

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Send Email'


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'gmail-python-email-send.json')
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


def SendMessage(sender, to, subject, msgHtml, msgPlain, attachment=None):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    if not attachment:
        message1 = CreateMessage(sender, to, subject, msgHtml, msgPlain)
    else:
        message1 = createMessageWithAttachment(sender, to, subject, msgHtml, msgPlain, attachment)
    SendMessageInternal(service, "me", message1)


def SendMessageInternal(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        # print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def CreateMessage(sender, to, subject, msgHtml, msgPlain):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to
    msg.attach(MIMEText(msgPlain, 'plain'))
    msg.attach(MIMEText(msgHtml, 'html'))
    raw = base64.urlsafe_b64encode(msg.as_bytes())
    raw = raw.decode()
    body = {'raw': raw}
    return body


def createMessageWithAttachment(
        sender, to, subject, msgHtml, msgPlain, attachmentFile):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      msgHtml: Html message to be sent
      msgPlain: Alternative plain text message for older email clients
      attachmentFile: The path to the file to be attached.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart('mixed')
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    messageA = MIMEMultipart('alternative')
    messageR = MIMEMultipart('related')

    messageR.attach(MIMEText(msgHtml, 'html'))
    messageA.attach(MIMEText(msgPlain, 'plain'))
    messageA.attach(messageR)

    message.attach(messageA)

    print("create_message_with_attachment: file:", attachmentFile)
    content_type, encoding = mimetypes.guess_type(attachmentFile)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(attachmentFile, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(attachmentFile, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(attachmentFile, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(attachmentFile, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(attachmentFile)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    message_as_bytes = message.as_bytes() # the message should converted from string to bytes.
    message_as_base64 = base64.urlsafe_b64encode(message_as_bytes) #encode in base64 (printable letters coding)
    raw = message_as_base64.decode()  # need to JSON serializable (no idea what does it means)
    return {'raw': raw}


def get_data_and_plot(fund_code_and_name, start_date, end_date):
    CORONATION_FUND_URL = 'www.coronation.com/FundPricesDownload'

    excel_filename = 'text.xlsx'

    n_funds = len(fund_code_and_name)
    fig = plt.figure(figsize=(18, 4.5*n_funds))
    gs = mpl.gridspec.GridSpec(n_funds, 2, width_ratios=[3, 1])

    for ii, fund in enumerate(fund_code_and_name):
        print('Getting data for fund {}'.format(fund))
        full_req = 'http://{0}/{1}/{2}/{3}'.format(CORONATION_FUND_URL, start_date, end_date, fund)
        resp = requests.get(full_req)

        with open(excel_filename, 'wb') as output:
            output.write(resp.content)

        df = pd.read_excel(excel_filename)

        # figure out what row data starts at by looking at the fund price column
        fund_prices = pd.to_numeric(df.iloc[:, 1], errors='coerce')

        non_nan_rows = fund_prices.notnull()
        fund_prices = fund_prices[non_nan_rows][::-1]
        dates = pd.to_datetime(df.iloc[:, 0][non_nan_rows])[::-1]

        centre_std_11 = fund_prices.rolling(window=11, center=True).std().fillna(0)
        centre_std_31 = fund_prices.rolling(window=31, center=True).std().fillna(0)

        df_fund = pd.DataFrame({
            'price': fund_prices,
            'date': dates,
            'p+std11': fund_prices+centre_std_11,
            'p-std11': fund_prices-centre_std_11,
            'p+std31': fund_prices+centre_std_31,
            'p-std31': fund_prices-centre_std_31
        })
        df_fund.set_index('date', inplace=True)

        # when should we have bought?
        regular_decrease = np.zeros_like(fund_prices, dtype=np.bool)
        decrease_days = 6
        for i in range(decrease_days, len(fund_prices)):
            regular_decrease[i] = np.all(fund_prices.iloc[i] < fund_prices.iloc[i-decrease_days:i])
        buy = True if np.any(regular_decrease[-3:]) else False

        ax0 = plt.subplot(gs[ii, 0])
        ax0.plot(df_fund.index[regular_decrease], fund_prices[regular_decrease], 'ro')
        ax0.plot(df_fund.index, fund_prices, color='#3F7BB9', marker='.', ls='-')
        ax0.fill_between(df_fund.index, df_fund['p-std11'], df_fund['p+std11'], facecolor='grey', alpha=0.4)
        ax0.fill_between(df_fund.index, df_fund['p-std31'], df_fund['p+std31'], facecolor='grey', alpha=0.3)
        ax0.set_title(fund_code_and_name[fund])
        #ax0.set_xlabel('Date')
        ax0.set_ylabel('Price / [ZAR]')
        ax0.xaxis.grid(True, ls='--')
        ax0.yaxis.grid(True, ls='--')

        # recent lookback plot
        lookback_days = 30
        lookback_std = fund_prices.rolling(window=15).std().fillna(0)[-lookback_days:]
        lookback_price = fund_prices[-lookback_days:]
        lookback_dates = df_fund.index[-lookback_days:]
        lookback_decrease = regular_decrease[-lookback_days:]

        ax1 = plt.subplot(gs[ii, 1])
        ax1.plot(lookback_dates[lookback_decrease], lookback_price[lookback_decrease], 'ro')
        ax1.plot(lookback_dates, lookback_price, color='#3F7BB9', marker='.', ls='-')
        ax1.fill_between(
            lookback_dates,
            lookback_price - lookback_std,
            lookback_price + lookback_std,
            facecolor='grey',
            alpha=0.4)
        #ax1.set_xlabel('Date')
        ax1.xaxis.grid(True, ls='--')
        ax1.yaxis.grid(True, ls='--')
        ax1.xaxis.set_major_locator(
            mpl.dates.WeekdayLocator(byweekday=mpl.dates.MO)
        )
        ax1.xaxis.set_major_formatter(
            mpl.dates.DateFormatter('%d %b\n%Y')
        )

        os.remove(excel_filename)

    figure_file = 'Coronation_fund_prices_{}.png'.format(end_date)
    plt.savefig('Coronation_fund_prices_{}.png'.format(end_date))
    plt.tight_layout()
    plt.show()

    return figure_file, buy


def main():

    fund_code_and_name = {
        'UTTOP': 'Top20',
        'UTINTG': 'Global Opportunities Equity [ZAR] Feeder',
        'UTCAPP': 'Capital Plus',
        'CIHEMF_USD_RETL': 'Global Emerging Markets',
        'CGESMF': 'Global Equity Select',
        'UTGESF': 'Global Equity Select [ZAR] Feeder',
        'UTSPGR': 'Smaller Companies',
        'UTGEMF': 'Global Emerging Markets Flexible [ZAR]'
    }

    start_date = '01-02-2016'
    now = pd.Timestamp(time.time(), unit='s')
    end_date = now.strftime('%d-%m-%Y')
    print('Date: ', end_date)
    figfile, buy = get_data_and_plot(fund_code_and_name, start_date, end_date)
    if buy:
        should_buy = ' - Time to buy?'
    else:
        should_buy = ''

    to = 'llrichter@gmail.com'
    sender = 'artisanaldata@gmail.com'
    subject = "{} - Coronation Unit Trust Prices {}".format(end_date, should_buy)
    msgHtml = "Hi Laura<br/>Here are your Coronation fund prices."
    msgPlain = "Hi\nHere are your Coronation fund prices."
    SendMessage(sender, to, subject, msgHtml, msgPlain, figfile)
    print()

if __name__ == '__main__':
    main()
