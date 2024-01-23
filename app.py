from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import os
import pandas as pd
import requests
from io import BytesIO

url = "https://raw.githubusercontent.com/dkgin/fate/main/三命通會.xlsx"
response = requests.get(url)
response.raise_for_status()

excel_data = pd.read_excel(BytesIO(response.content), engine='openpyxl')

app = Flask(__name__)

@app.route("/callback", methods=['POST'])
def callback():
    body = request.get_data(as_text=True)                    # 取得收到的訊息內容
    try:
        json_data = json.loads(body)                         # json 格式化訊息內容
        line_bot_api = LineBotApi(os.environ['CHANNEL_ACCESS_TOKEN'])
        handler = WebhookHandler(os.environ['CHANNEL_SECRET'])
        signature = request.headers['X-Line-Signature']      # 加入回傳的 headers
        handler.handle(body, signature)                      # 綁定訊息回傳的相關資訊
        msg = json_data['events'][0]['message']['text']      # 取得 LINE 收到的文字訊息
        tk = json_data['events'][0]['replyToken']            # 取得回傳訊息的 Token
        response = get_response_from_excel(msg)
        line_bot_api.reply_message(tk,TextSendMessage(text=response))  # 回傳訊息
        print(msg, tk)                                       # 印出內容
    except Exception as e:
        print(f"Error: {e}")
        print(body)                                          # 如果發生錯誤，印出收到的內容
    return 'OK'                 # 驗證 Webhook 使用，不能省略

import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

def get_response_from_excel(input_text):
    try:
        results_str = f'{input_text} 年次的訥音\n'
        for  sheet_name, df in excel_data.items():
            result = df[df["農曆年次"] == input_text]
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