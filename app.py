from pandas import read_excel, DataFrame
from yaml import reader

# https://github.com/ghpranav/python-bulk-mail/blob/master/mail.py


def sendMail():
    data :DataFrame = read_excel('example-data.xlsx')
    print(data.iloc[0].loc['Email'])



if __name__ == '__main__':
    sendMail()