from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
import os
import pandas as pd
from linebot.models import MessageEvent, TextMessage, TextSendMessage, StickerSendMessage, ImageSendMessage, LocationSendMessage
import requests, json, time, statistics
from io import BytesIO

url = "https://raw.githubusercontent.com/dkgin/fate/main/三命通會.xlsx"
response = requests.get(url)
response.raise_for_status()

excel_data = pd.read_excel(BytesIO(response.content), engine='openpyxl')

# LINE 回傳訊息函式
def reply_message(msg, rk, token):
  headers = {'Authorization':f'Bearer {token}','Content-Type':'application/json'}
  body = {
  'replyToken':rk,
  'messages':[{
      "type": "text",
      "text": msg
      }]
  }
  req = requests.request('POST', 'https://api.line.me/v2/bot/message/reply', headers=headers,data=json.dumps(body).encode('utf-8'))
  print(req.text)

# 三命通會
def fate(msg):
    try:
        results_str = f'{msg} 年次的訥音\n'
        for sheet_name, df in excel_data.items():
            result = df[df["農曆年次"] == msg]
            if not result.empty:
                calculate = result["天干地支"].values[0]
                attribute = result["訥音"].values[0]
                good = result["吉訥音"].values[0]
                bad = result["凶訥音"].values[0]
                for index, row in result.iterrows():
                  results_str += f'【{row["天干地支"]}年，{row["訥音"]}】\n\n 【吉】{row["吉訥音"]}\n\n 【凶】{row["凶訥音"]}\n\n 【論訥音】{row["論訥音吉凶"]}\n'
                return results_str
        return "No matching response found"
    except Exception as e:
        return f"Error: {str(e)}"

app = Flask(__name__)

access_token = '0VwSNMKzkkxvqbBG0QJHjOro/LAbii4Fmn47oG3Dnjnj3ZVDUu+o8a0ZqMIDHNI0ZdZsVKthhCDJhsDN+hEUEeE0pT6qyhQmDoeoYwhI8Vwy3/vZdpKUEYE7YvF5Yf2niVZ+6xT9hTvoTT5PcXHm8QdB04t89/1O/w1cDnyilFU='

@app.route("/callback", methods=['POST'])
def callback():
    body = request.get_data(as_text=True)                    # 取得收到的訊息內容
    try:
        json_data = json.loads(body)                         # json 格式化訊息內容
        line_bot_api = LineBotApi(os.environ['CHANNEL_ACCESS_TOKEN'])              # 確認 token 是否正確
        handler = WebhookHandler(os.environ['CHANNEL_SECRET'])                     # 確認 secret 是否正確
        signature = request.headers['X-Line-Signature']      # 加入回傳的 headers
        handler.handle(body, signature)                      # 綁定訊息回傳的相關資訊
        if 'events' in json_data and json_data['events']:
            reply_token = json_data['events'][0]['replyToken']    # 取得回傳訊息的 Token ( reply message 使用 )
            user_id = json_data['events'][0]['source']['userId']  # 取得使用者 ID ( push message 使用 )
            msg_type = json_data['events'][0]['message']['type']
            print(json_data)
            if msg_type == 'text':
              msg = json_data['events'][0]['message']['text']
              reply_message(f'{fate(msg)}', reply_token, access_token)

    except InvalidSignatureError:
      abort(400)
    return 'OK'                 # 驗證 Webhook 使用，不能省略

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
