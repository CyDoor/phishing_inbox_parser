__author__ = 'F9XK3LI'

import time
import win32com.client
import os
from CONFIG import *

OUTLOOK = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')

def get_emails_from_shared():

    namespace = OUTLOOK.Session
    recipient = namespace.CreateRecipient(SHARED_INBOX)
    inbox = OUTLOOK.GetSharedDefaultFolder(recipient, 6)
    messages = inbox.Items
    for message in messages:
        try:
            get_attachments(message)
        except:
            print "Failed for ", message


def get_attachments(message):
    attachments = message.Attachments
    for attachment in attachments:
        print attachment.FileName
        attachment.SaveASFile(
            os.path.dirname(os.path.abspath(__file__)) + '/attachments/' + str(time.time()) + attachment.FileName)


def parse_email(file_path):
    msg = OUTLOOK.OpenSharedItem(file_path)
    return msg


def get_urls(message):
    pass

def get_header_and_attachments(message):
    dict = {}
    dict["TransportMessageHeaders"] = message.PropertyAccessor()
    dict["SenderName"] = message.SenderName
    dict["SenderEmailAddress"] = message.SenderEmailAddress
    dict["SentOn"] = message.SentOn
    dict["To"] = message.To
    dict["CC"] = message.CC
    dict["BCC"] = message.BCC
    dict["Subject"] = message.Subject
    dict["Body"] = message.Body
    attachments = message.Attachments
    dict["Attachments"] = []
    for attachment in attachments:
        dict["Attachments"].append(attachment.FileName)
    return dict


def absolute_paths(directory):
    for dirpath, _, filenames in os.walk(directory):
        for f in filenames:
            yield os.path.abspath(os.path.join(dirpath, f))


if __name__ == "__main__":
    #get_emails_from_shared()
    for file_path in absolute_paths("attachments"):
        print get_header_and_attachments(parse_email(file_path))