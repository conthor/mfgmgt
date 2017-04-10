import os
from slackclient import SlackClient
import time
import smartsheet

def list_channels():
    channels_call = slack_client.api_call("channels.list")
    if channels_call.get('ok'):
        return channels_call['channels']
    return None

def channel_info(channel_id):
    channel_info = slack_client.api_call("channels.info", channel=channel_id)
    if channel_info:
        return channel_info['channel']
    return None

def send_message(channel_id, message):
    slack_client.api_call(
        "chat.postMessage",
        channel=channel_id,
        text=message,
        username='ptrnotifier',
        icon_emoji=':robot_face:'
    )



sc = SlackClient('xoxp')
CHANNEL_NAME = "barcode"
sc.rtm_connect()
xSheet = smartsheet.Smartsheet('lswl')
# Class object
xResult = xSheet.Sheets.list_sheets(include_all=True)
# The list
xList = []
for sheet1 in xResult.data:
    xList.append((sheet1._name.encode('utf-8'),sheet1._modified_at))
# sort the list created by 'Modifiedat' attribute
xNlist = sorted(xList,key=lambda x: x[1])
# print list
for key, value in xNlist:
    print key,value
#slack_client.api_call("api.test")
# send list
for key, value in xNlist:
        sc.rtm_send_message("barcode", key)
while True:
            # Read latest messages
            for slack_message in sc.rtm_read():
                message = slack_message.get("text")
                print "Message",message
                user = slack_message.get("user")
                if not message or not user:
                    continue
                #sc.rtm_send_message(CHANNEL_NAME, "<@{}> send to Smartsheet ".format(user))
                #sc.rtm_send_message(CHANNEL_NAME, message)
                
            # Sleep for half a second
            time.sleep(0.5)

if __name__ == '__main__':
    channels = list_channels()
    if channels:
        print("Channels: ")
        for channel in channels:
            print(channel['name'] + " (" + channel['id'] + ")")
            detailed_info = channel_info(channel['id'])
            if detailed_info:
                print('Latest text from ' + channel['name'] + ":")
                #print(detailed_info['latest']['text'])
            if channel['name'] == 'general':
                send_message(channel['id'], "Hello " +
                             channel['name'] + "! It worked!")
        print('-----')
    else:
        print("Unable to authenticate.")
