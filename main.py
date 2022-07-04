# coding=utf-8
import argparse
import clr
import ConfigParser
import ctypes

from WebLibrary import WebLibrary
from keyring import get_password

parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group(required=True)
group.add_argument("-i", "--clock-in", action="store_true", help="clocks you in Qandle")
group.add_argument("-o", "--clock-out", action="store_true", help="clocks you out of Qandle")
args = parser.parse_args()


class Qandle:
    def __init__(self):
        self.user32 = ctypes.windll.user32
        self.clocked = False

        self.config = ConfigParser.RawConfigParser()
        self.config.read("config.cfg")

        self.__get_credential()

        self.web = WebLibrary()

        self.web.open_browser(self.config.get("sign in", "url"))
        self.web.wait_until_element_is_visible(self.config.get("sign in", "email"))

    def __get_credential(self):
        password = get_password(self.config.get("sign in", "url"), self.config.get("sign in", "username"))
        return password

    def __message(self, message):
        self.user32.MessageBoxW(None, message, u"Qandle", 0)

    def login(self):
        self.web.input_text_by_xpath(self.config.get("sign in", "email"), self.config.get("sign in", "username"))
        self.web.input_password_by_xpath(self.config.get("sign in", "password"), self.__get_credential())
        self.web.click_element_by_xpath(self.config.get("sign in", "signin"))
        self.web.wait_until_element_is_visible(self.config.get("left pane", "username"))

    def clock_in(self):
        self.web.wait_until_element_is_visible(self.config.get("clock tile", "clock-in"))
        self.web.click_element_by_xpath(self.config.get("clock tile", "clock-in"))

        # self.__message("User successfully clocked in")
        self.clocked = True

    def clock_out(self):
        self.web.wait_until_element_is_visible(self.config.get("clock tile", "clock-out"))
        self.web.click_element_by_xpath(self.config.get("clock tile", "clock-out"))

        self.web.wait_until_element_is_visible(self.config.get("clock tile", "clock-out_confirmation"))
        self.web.click_element_by_xpath(self.config.get("clock tile", "clock-out_confirmation"))

        # self.__message("User successfully clocked out")
        self.clocked = True

    def log_out(self):
        if not self.clocked:
            action = "in" if args.clock_in else "out"
            self.__message(u"User have already clocked {}".format(action))

        try:
            self.web.wait_until_element_is_visible(self.config.get("logout", "logout_arrow"))
            self.web.click_element_by_xpath(self.config.get("logout", "logout_arrow"))

            self.web.wait_until_element_is_visible(self.config.get("logout", "logout"))
            self.web.click_element_by_xpath(self.config.get("logout", "logout"))
        finally:
            self.web.close_browser()



if __name__ == "__main__":
    qandle = Qandle()
    qandle.login()
    try:
        if args.clock_in:
            qandle.clock_in()
        elif args.clock_out:
            qandle.clock_out()
    finally:
        qandle.log_out()
