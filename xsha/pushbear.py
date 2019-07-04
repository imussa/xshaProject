# Server酱 http://sc.ftqq.com
import requests


class WeMSG:
    SCKEY = "SCU54202Tbe18f7a69ff78c24d1c2f3cf3bae058a5d12f33cc2cb4"
    URL = "https://sc.ftqq.com/{SCKEY}.send".format(SCKEY=SCKEY)

    def __init__(self, title="标题", message=""):
        self.title = title
        if isinstance(message, str):
            self.message = message
        else:
            self.message = message.read()
        self._push()

    def _push(self):
        params = {"text":self.title, "desp":self.message}
        r = requests.get(url=self.URL, params=params)
        self._response = r.json()

    @property
    def response(self):
        return self._response['errmsg']


if __name__ == "__main__":
    x = WeMSG()
    print(x.response)
