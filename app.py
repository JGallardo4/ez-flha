#!/usr/bin/python

import click
import config
from imap_tools import MailBox

@click.command()
def process_flhas():
    """ This program downloads pdf files from email, converts to PNG, and submits to Sitedocs. """

    with MailBox(config.SMTP_SERVER).login(config.FROM_EMAIL, config.APP_PWD, 'INBOX') as mailbox:
        for msg in mailbox.fetch():
            for att in msg.attachments:
                print(att.filename, att.content_type)


if __name__ == '__main__':
    process_flhas()
