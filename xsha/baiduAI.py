from aip import AipSpeech


class Speech:

    def __init__(self):
        APP_ID = '11468048'
        API_KEY = 'BF0OERtklEYDGG9k6mGh5cL4'
        SECRET_KEY = 'alw8VKwzadfbWfP5T1iM6xVjxFimfLXd'
        self.client = AipSpeech(APP_ID, API_KEY, SECRET_KEY)

    def get_speech(self, text, file=None, lang='zh', ctp=1, spd=5, pit=5, vol=5, per=0):
        """

        :param text: 合成语音的文本
        :param file: 文件流
        :param lang: 合成语种
        :param ctp: 默认为1
        :param spd: 语速，取值0-9，默认为5中语速
        :param pit: 音调，取值0-9，默认为5中语调
        :param vol: 音量，取值0-15，默认为5中音量
        :param per: 发音人选择, 0为女声，1为男声，3为情感合成-度逍遥，4为情感合成-度丫丫，默认为普通女
        :return: None
        """
        options = {"spd": spd, "pit": pit, "vol": vol, "per": per}
        result = self.client.synthesis(text, lang, ctp, options)
        if file and not isinstance(result, dict):
            file.write(result)
        else:
            return result


if __name__ == "__main__":
    sp = Speech()
    text = "这是一条测试！"
    with open("audio.mp3", "wb") as f:
        sp.get_speech(text, file=f)

