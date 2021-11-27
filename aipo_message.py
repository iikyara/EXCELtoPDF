#! /usr/bin/python
# coding: utf-8

import sys
import os
import re
import requests

## settings
aipo_url = ''
username = ''
password = ''

## constants
headers = {'User-Agent': 'Mozilla/5.0'} # UAが python-requests だと Aipo がバグる
portlet_id = ''

def save_aipouserinfo(url, usrnm, pwd, p_id):
    global aipo_url
    global username
    global password
    global portlet_id
    aipo_url = url
    username = usrnm
    password = pwd
    portlet_id = p_id
    print(
        "aipo_url : %s\n" % aipo_url,
        "username : %s\n" % username,
        "password : %s\n" % password,
        "portlet_id : %s\n" % portlet_id
    )

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
