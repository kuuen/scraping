import PySimpleGUI as sg
import AmazonScraping
import YahooScraping
from logging import getLogger, StreamHandler, DEBUG
import logging
import sys
logger = getLogger(__name__)
handler = StreamHandler()
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False
log_file = logging.FileHandler('log/app.log')
logger.addHandler(log_file)

# pip install PySimpleGUI

layout = [  
    [sg.Text('対象サイト'), sg.Combo(['', 'Yahooショッピング', 'Amazon'], key='site', size=(20, 15))],
    [sg.Text('キーワード'), sg.Input('', key = 'keyword')],
    [sg.Button('実行', key='exec'),  sg.Button('終了', key='end')]]

window = sg.Window('店舗登録', layout)

while True:
  event, values = window.read()
  if event == sg.WIN_CLOSED or event == 'end':
    break
  elif event == 'exec':
    errorlist = list()
    if values['site'] == '':
        errorlist.append ('対象サイトを指定してください')
    if values['keyword'] == '':
        errorlist.append('キーワードを指定してください')

    if len(errorlist) > 0:
      msg = ''
      for error in errorlist:
        msg += error + '\n'

      # 最後の改行を取り除く
      msg = msg[0 : len(msg) - 1]

      sg.popup_error(msg)
    else :
      try :
        if values['site'] == 'Yahooショッピング':
          rtnmsg = YahooScraping.main(values['keyword'])
        elif values['site'] == 'Amazon':
          rtnmsg = AmazonScraping.main(values['keyword'])

        sg.popup_ok(rtnmsg)
      except Exception as e:
        import traceback
        logger.error("サイト: %s キーワード : %s" % (values['site'], values['keyword']))
        logger.error(traceback.format_exc())
        
window.close()