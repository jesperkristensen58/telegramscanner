"""
Parse messages from Telegram into local Excel Workbook.

@author Jesper Kristensen
"""
import datetime
import json
from xlwt import Workbook
from collections import defaultdict

with open("channel_messages.json", "r") as fd:
    messages = json.load(fd)

# Workbook is created
wb = Workbook()
# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Twitter Profiles')

sheet1.write(1, 0, 'Date:')
sheet1.write(1, 1, 'Name:')
sheet1.write(1, 2, 'Handle:')
sheet1.write(1, 3, 'Url:')
sheet1.write(1, 4, 'Description:')

handles_already_seen = set()

rowix = 0
profiles = []
duplicate_profiles = defaultdict(int)
for msg in messages:

    if "message" not in msg:
        continue

    txt = msg['message']
    if "twitter" in txt.lower():
        try:
            _ = msg['media']['webpage']['site_name']
        except:
            continue

        if msg['media']['webpage']['site_name'].lower().strip() == 'twitter':

            date = datetime.datetime.strptime(msg["date"], "%Y-%m-%dT%H:%M:%S+00:00").strftime("%d-%b-%Y %H:%M:%S")
            name = msg['media']['webpage']['title']
            descr = msg['media']['webpage']['description']
            handle = msg['media']['webpage']['display_url'][msg['media']['webpage']['display_url'].find('/') + 1:]
            url = msg["media"]["webpage"]["url"]

            if handle in handles_already_seen:
                duplicate_profiles[handle] += 1
                continue

            profile = {"date": date,
                       "name": name,
                       "descr": descr,
                       "handle": handle,
                       "url": url}

            profiles.append(profile)

            sheet1.write(2 + rowix, 0, date)
            sheet1.write(2 + rowix, 1, name)
            sheet1.write(2 + rowix, 2, handle)
            sheet1.write(2 + rowix, 3, url)
            sheet1.write(2 + rowix, 4, descr)

            rowix += 1

            handles_already_seen.add(handle)


wb.save(f"TwitterProfiles-{datetime.datetime.now().strftime('%b-%d-%y')}.xls")


print("Duplicate profiles posted in the channel:")

duplicate_profiles = {k: v for k, v in sorted(duplicate_profiles.items(), key=lambda item: -item[1])}

for pf in duplicate_profiles:
    print(f"{pf} posted {duplicate_profiles[pf] + 1} times.")
