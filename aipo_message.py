#! /usr/bin/python
# coding: utf-8

import sys
import os
import re
import requests

## settings
aipo_url = 'https://portal.coopii.com/aipo/'
username = 'sugi'
password = 'sugiyan2543'

'''
month = '4'
members = [
    {'room_id': '4', 'name': '小野田　春希'},
    {'room_id': '10', 'name': '村野　直樹'},
    {'room_id': '11', 'name': '木村　隆秀'},
    {'room_id': '12', 'name': '小島　馨'},
    {'room_id': '13', 'name': '後藤　勇光'},
    {'room_id': '14', 'name': '齋藤　勲'},
    {'room_id': '15', 'name': '坂田　直樹'},
    {'room_id': '16', 'name': '櫻井　直輝'},
    {'room_id': '17', 'name': '清藤　雄一'},
    {'room_id': '18', 'name': '寺内　綾香'},
    {'room_id': '19', 'name': '松村　英則'},
    {'room_id': '20', 'name': '山下　雅史'},
]
message_template = '%MONTH%月度給与明細です。\nお疲れ様でした。'
attachment_template = '/tmp/%NAME%給与（%MONTH%月）明細.zip'
##
'''

## constants
headers = {'User-Agent': 'Mozilla/5.0'} # UAが python-requests だと Aipo がバグる
portlet_id = 'P-14b386ab648-100f9'


def get_aipo_session():
    response = requests.get(aipo_url, headers=headers)
    if not 'JSESSIONID' in response.cookies:
        print('invalid aipo url')
    else:
        return response.cookies['JSESSIONID']

def aipo_login(jsessionid, username, password):
    cookies = {'JSESSIONID': jsessionid}
    login_data = {'action': 'ALJLoginUser', 'username': username, 'password': password}
    response = requests.post(aipo_url, data=login_data, headers=headers, cookies=cookies)
    if not 'JSESSIONID' in response.cookies:
        print('invalid username/password')
    else:
        return response.cookies['JSESSIONID']

def get_aipo_messageform(jsessionid):
    cookies = {'JSESSIONID': jsessionid}
    params = {'template': 'MessageFileuploadFormScreen', 'mode': 'miniform'}
    response = requests.get(aipo_url + 'portal/js_peid/global-' + portlet_id, params=params, headers=headers, cookies=cookies)
    match_f = re.search('id="folderName" value="([^"]+)"', response.text)
    match_s = re.search('name="secid" value="([^"]+)"', response.text)
    if match_f and match_s:
        foldername = match_f.group(1)
        secid = match_s.group(1)
        return foldername, secid
    else:
        print('cannot get message form')

def put_aipo_message_attachment(jsessionid, foldername, secid, attachment_file):
    attachment_content = get_content(attachment_file)
    if attachment_content:
        cookies = {'JSESSIONID': jsessionid}
        upload_params = {'template': 'MessageFileuploadFormScreen', 'mode': 'update-mini', 'msize': '0'}
        upload_data = {'mode': 'upload', 'folderName': foldername, 'msize': '0',  'nsize': '0', 'secid': secid}
        upload_files = {'attachment': (os.path.basename(attachment_file), attachment_content, 'application/octet-stream')}
        response = requests.post(aipo_url + 'portal/js_peid/global-' + portlet_id, data=upload_data, params=upload_params, files=upload_files, headers=headers, cookies=cookies)
        match = re.search("window.parent.aipo.fileupload.onAddFileInfo\('(\d_\d+)','([-\d]+)'", response.text)
        if match:
            return match.group(2)
        else:
            print('upload fileid not found')

def put_aipo_message(jsessionid, room_id, foldername, secid, attachment_id, message):
    cookies = {'JSESSIONID': jsessionid}
    message_params = {'template': 'MessageFormJSONScreen'}
    message_data = {'_name': 'formMessage', 'roomId': room_id, 'userId': '0', 'secid': secid, 'folderName': foldername, 'attachments': attachment_id, 'message': message}
    response = requests.post(aipo_url + 'portal/js_peid/' + portlet_id, data=message_data, params=message_params, headers=headers, cookies=cookies)
    match = re.search('err', response.text)
    if match:
        print('post error')
    else:
        return True

def post_message(jsessionid, room_id, attachment_file, message):
    form_param = get_aipo_messageform(jsessionid)
    if not form_param:
        return None
    foldername, secid = form_param[0], form_param[1]
    attachment_id = put_aipo_message_attachment(jsessionid, foldername, secid, attachment_file)
    if not attachment_id:
        return None
    return put_aipo_message(jsessionid, room_id, foldername, secid, attachment_id, message)

def get_content(attachment_file):
    try:
        with open(attachment_file, mode="rb") as file:
            attachment_content = file.read()
    except IOError:
        print('cannot open attachment file')
    else:
        return attachment_content

# 処理ここから
if __name__ == "__main__":
    jsessionid = get_aipo_session()
    if not jsessionid:
        sys.exit(1)

    jsessionid = aipo_login(jsessionid, username, password)
    if not jsessionid:
        sys.exit(1)

    for member in members:
        attachment_file = attachment_template.replace('%NAME%', member['name']).replace('%MONTH%', month)
        message = message_template.replace('%NAME%', member['name']).replace('%MONTH%', month)
        if not post_message(jsessionid, member['room_id'], attachment_file, message):
            print('post error on ' + member['name'])
