#!/usr/bin python3

import mailbox
import os
import sys
import traceback
from os import getenv
from dotenv import load_dotenv
from exchangelib import IMPERSONATION, Account, Credentials, OAuth2Credentials, \
    Configuration, OAUTH2, Identity

load_dotenv()

client_id=getenv('client_id')
client_secret=getenv('client_secret')
tenant_id=getenv('tenant_id')
user=getenv('user_mail')

ID_FILE = '.read_ids'


def create_mailbox_message(e_msg):
    m = mailbox.mboxMessage(e_msg.mime_content)
    if e_msg.is_read:
        m.set_flags('S')
    return m

def get_read_ids():
    if os.path.exists(ID_FILE):
        with open(ID_FILE, 'r') as f:
            return set([s for s in f.read().splitlines() if s])
    else:
        return set()

def set_read_ids(ids):
    with open(ID_FILE, 'w') as f:
        for i in ids:
            if i:
                f.write(i)
                f.write(os.linesep)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: {} folder_name mbox_file".format(sys.argv[0]))
        sys.exit()
    credentials = OAuth2Credentials(
        client_id=client_id,
        client_secret=client_secret,
        tenant_id=tenant_id,
        identity=Identity(primary_smtp_address=user),
    )
    config = Configuration(
        credentials=credentials,
        auth_type=OAUTH2,
        service_endpoint="https://outlook.office365.com/EWS/Exchange.asmx",
    )
    account = Account(
        user,
        config=config,
        autodiscover=False,
        access_type=IMPERSONATION,
    )
    mbox = mailbox.mbox(sys.argv[2])
    mbox.lock()
    read_ids_local = get_read_ids()
    folder = getattr(account, sys.argv[1], None)
    item_ids_remote = list(folder.all().order_by('-datetime_received').values_list('id', 'changekey'))
    total_items_remote = len(item_ids_remote)
    new_ids = [x for x in item_ids_remote if x[0] not in read_ids_local]
    read_ids = set()
    print("Total items in folder {}: {}".format(sys.argv[1], total_items_remote))
    for i, item in enumerate(account.fetch(new_ids), 1):
        try:
            msg = create_mailbox_message(item)
            mbox.add(msg)
            mbox.flush()
        except Exception as e:
            traceback.print_exc()
            print("[ERROR] {} {}".format(item.datetime_received, item.subject))
        else:
            if item.id:
                read_ids.add(item.id)
            print("[{}/{}] {} {}".format(i, len(new_ids), str(item.datetime_received), item.subject))
    mbox.unlock()
    set_read_ids(read_ids_local | read_ids)