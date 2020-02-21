from lxml import etree
from requests import get


class Ishadowx:
    url = 'http://isx.yt/'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.87 Safari/537.36'}

    @staticmethod
    def get_selector():
        resp = get(Ishadowx.url, headers=Ishadowx.headers)
        selector = etree.HTML(resp.content)
        return selector


class ServerConf:
    _ip_xpath = "//div[@class='hover-text']/h4/span[@id='ip{server}']/text()"
    _port_xpath = "//div[@class='hover-text']/h4/span[@id='port{server}']/text()"
    _pw_xpath = "//div[@class='hover-text']/h4/span[@id='pw{server}']/text()"

    @staticmethod
    def factory():
        server_list = ['usa', 'usb', 'usc', 'sga', 'sgb', 'sgc']
        return list(map(ServerConf, server_list))

    def __init__(self, server):
        self.ip = self.port = self.pw = server

    @property
    def ip(self):
        return selector.xpath(self._ip_xpath)[0].strip()

    @ip.setter
    def ip(self, value):
        self._ip_xpath = self._ip_xpath.format(server=value)

    @property
    def port(self):
        return selector.xpath(self._port_xpath)[0].strip()

    @port.setter
    def port(self, value):
        self._port_xpath = self._port_xpath.format(server=value)

    @property
    def pw(self):
        return selector.xpath(self._pw_xpath)[0].strip()

    @pw.setter
    def pw(self, value):
        self._pw_xpath = self._pw_xpath.format(server=value)


if __name__ == "__main__":
    selector = Ishadowx.get_selector()
    server = ServerConf.factory()
    print(server)
