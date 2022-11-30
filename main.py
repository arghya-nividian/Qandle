# coding=utf-8
import argparse
import os
import sys
import time

import clr
try:
    import ConfigParser
except ImportError:
    import configparser as ConfigParser
import ctypes
import logging
from selenium.common.exceptions import WebDriverException

from WebLibrary import WebLibrary
from keyring import get_password

parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group(required=True)
group.add_argument("-i", "--clock-in", action="store_true", help="clocks you in Qandle")
group.add_argument("-o", "--clock-out", action="store_true", help="clocks you out of Qandle")
args = parser.parse_args()

logging.basicConfig(format="%(asctime)s - %(levelname)-10s - %(message)s")
logging.addLevelName(logging.INFO + 5, "SUCCESS")
log = logging.getLogger("root")
log.setLevel(logging.INFO)

log.success = lambda message: log._log(logging.INFO + 5, message, ())


class Qandle:
    def __init__(self):
        self.user32 = ctypes.windll.user32
        self.clocked = False

        self.config = ConfigParser.RawConfigParser()
        self.config.read("config.cfg")

        self.web = WebLibrary()

        try:
            log.info("Launching browser and opening Qandle url")
            self.web.open_browser(self.config.get("sign in", "url"))
        except WebDriverException:
            log.warning("Device is not connected to internet. Closing the browser")
            self.web.close_browser()
            log.info("Browser is closed. Waiting for the user to fix the connectivity issue.")
            self.message(u"please fix your internet connection, and press 'OK'")
            log.info("User response received.")
            raise
        else:
            try:
                self.web.wait_until_element_is_visible(self.config.get("sign in", "email"))
                log.info("Qandle is loaded")
            except AssertionError:
                log.warning("Device is not connected to internet. Closing the browser")
                self.web.close_browser()
                log.info("Browser is closed. Waiting for the user to fix the connectivity issue.")
                self.message(u"please fix your internet connection, and press 'OK'")
                log.info("User response received.")
                raise

    def __get_credential(self):
        password = get_password(self.config.get("sign in", "url"), self.config.get("sign in", "username"))
        return password

    def message(self, message):
        self.user32.MessageBoxW(None, message, u"Qandle", 0)

    def login(self):
        log.info("Entering 'Work Email'")
        self.web.input_text_by_xpath(self.config.get("sign in", "email"), self.config.get("sign in", "username"))
        log.info("Entering 'Work Email' - Success")
        log.info("Entering 'Password'")
        self.web.input_text_by_xpath(self.config.get("sign in", "password"), self.__get_credential())
        log.info("Entering 'Password' - Success")
        log.info("Clicking on 'SIGN IN'")
        self.web.click_element_by_xpath(self.config.get("sign in", "signin"))
        log.info("Clicking on 'SIGN IN' - Success")
        self.web.wait_until_element_is_visible(self.config.get("left pane", "username"))
        log.info("Logging into 'Qandle' - Success")

    def navigate_to_dashboard(self):
        for _ in range(3):
            try:
                self.web.wait_until_element_is_visible(self.config.get("clock tile", "title"))
                log.info("Navigating to 'Dashboard' - Success")
                return
            except AssertionError:
                log.warning("Navigating to 'Dashboard' - Failure - Reloading page")
                self.web.reload_page()
        else:
            raise AssertionError

    def clock_in(self):
        status = self.check_exists(timeout_in_second=10,
                                   clock_in=self.config.get("clock tile", "clock-in"),
                                   start_break=self.config.get("clock tile", "start-break")
                                   )
        if status == "clock_in":
            log.info("Clicking on 'Clock In'")
            self.web.wait_until_element_is_visible(self.config.get("clock tile", "clock-in"))
            self.web.click_element_by_xpath(self.config.get("clock tile", "clock-in"))
            log.info("Clicking on 'Clock In' - Success")

            # self.__message("User successfully clocked in")
            self.clocked = True
        else:
            log.warning("User has already clocked in")
            self.message(u"User has already clocked in")

    def clock_out(self):
        status = self.check_exists(timeout_in_second=10,
                                   clock_in=self.config.get("clock tile", "clock-in"),
                                   clock_out=self.config.get("clock tile", "clock-out")
                                   )
        if status == "clock_out":
            log.info("Clicking on 'Clock Out'")
            self.web.wait_until_element_is_visible(self.config.get("clock tile", "clock-out"))
            self.web.click_element_by_xpath(self.config.get("clock tile", "clock-out"))
            log.info("Clicking on 'Clock Out' - Success")

            log.info("Clicking on 'Yes'")
            self.web.wait_until_element_is_visible(self.config.get("clock tile", "clock-out_confirmation"))
            self.web.click_element_by_xpath(self.config.get("clock tile", "clock-out_confirmation"))
            log.info("Clicking on 'Yes' - Success")

            # self.__message("User successfully clocked out")
            self.clocked = True
        else:
            log.warning("User has already clocked out")
            self.message(u"User has already clocked out")

    def log_out(self):
        try:
            log.info("Expanding logout dropdown")
            self.web.wait_until_element_is_visible(self.config.get("logout", "logout_arrow"))
            self.web.click_element_by_xpath(self.config.get("logout", "logout_arrow"))
            log.info("Expanding logout dropdown - Success")

            log.info("Expanding logout")
            self.web.wait_until_element_is_visible(self.config.get("logout", "logout"))
            self.web.click_element_by_xpath(self.config.get("logout", "logout"))
            log.info("Expanding logout - Success")
        finally:
            log.info("Closing browser")
            self.web.close_browser()
            log.info("Closing browser - Success")

    def check_exists(self, timeout_in_second, **elements):
        """
        Waits for a set of xpaths to be available in the webpage, and returns the name of the xpaths which ever is found
        first.

        PARAMETERS
        ==========
        timeout_in_second: int
            number of seconds to wait before raising an exception
        elements: str
            named xpaths of the elements to wait for


        RETURNS
        =======
        str
            key_name of the xpath that was found first

        RAISES
        ======
        Exception
            none of the provided xpaths were found within the given timeout

        EXAMPLES
        ========
        >>> self.check_exists(120, home_page="xpath1", login_page="xpath2")
        'login_page'
        """
        for i in range(int(timeout_in_second)):
            for key, xpath in elements.items():
                if self.web.page_should_contain_element_by_xpath(xpath):
                    return key
                else:
                    continue
            else:
                time.sleep(1)
        else:
            raise Exception("None of the elements found within {} seconds : {}".format(timeout_in_second,
                                                                                       elements.keys()))


if __name__ == "__main__":
    log.warning(args)
    for _ in range(3):
        try:
            qandle = Qandle()
        except WebDriverException:
            time.sleep(10)
            continue
        else:
            qandle.login()

            qandle.navigate_to_dashboard()
            try:
                if args.clock_in:
                    qandle.clock_in()
                    log.success("Clocking in - Successful")
                elif args.clock_out:
                    qandle.clock_out()
                    log.success("Clocking out - Successful")
            finally:
                qandle.log_out()
                os.system("pause")
                sys.exit()
