# -*- coding: utf-8 -*-
import json
import requests


def send_msg(question):
    url = 'http://192.168.59.9:9996/api/v1/debug'  # http://192.168.59.9:9999/debug  http://113.207.31.77:10002/debug http://192.168.59.4:9996/api/v1/debug

    headers = {"content-type": "application/json"}
    question=question.replace("\n", "").replace(" ", "")
    data = {

        "tenant-id": 8,
        "chat_close": True,
        "net_close": True,
        "queue-type": "Q_QUEUE",
        "type": "get_answer_debug",
        "parameters": [{"question": question

                        }]}
    r = requests.post(url, data=json.dumps(data), headers=headers)
    result = r.text
    rest = json.loads(result)
    print 'question:' + question
    return rest


# if __name__ == "__main__":
#     question = '如何理解大数据/n'
#     xx = send_msg(question)
#     print xx
