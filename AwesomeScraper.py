# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium import common
from selenium.common.exceptions import *
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from bs4 import BeautifulSoup
import time
import urllib2
from threading import Thread
import Tkinter
from tkFileDialog import askopenfilename
import tkMessageBox
import Queue
import exceptions
import os
import webbrowser
import traceback
import logging
import subprocess
from win32process import *


__author__ = "Michael Yuan"
__copyright__ = "Copyright 2016"
__credits__ = "Catrina Meng"
__license__ = "GPL"
__version__ = "v1.0.7"


# Generate UTF-8 encoded url link of the search path.
def search_link_generator(entry_list):
    link_list = []
    domain = 'http://www.tianyancha.com/search?key='
    for entry in entry_list:
        info = entry.encode('utf-8')
        link = domain + urllib2.quote(info)
        link_list.append(link)
    return link_list


def write_file(info, path='result.txt'):
    if not os.path.exists('result'):
        os.makedirs('result')
    f = open('result\\' + path, 'a')
    try:
        f.write(info)
    finally:
        f.close()


class InfoScraper(object):
    def __init__(self, src_path, dst_path):
        self.src_path = src_path
        self.dst_path = dst_path
        self.load_timeout = 60
        self.browser = None
        self.js_subprocess = None
        self.restart()

        self.format_error = False           # Flag indicating format error of user specified file.
        self.company_list = []              # List read from user specified txt file.
        self.url_list = []                  # URL List generated from company_list.
        self.abort = False                  # Flag for signaling the abort of scraper and monitor thread.
        self.service_denied = False         # Flag indicating the current search url is denied by server.
        self.service_denied_count = 0       # Flag indicating how many times it has been denied.
        self.service_denied_limits = 4      # Maximum deny count.
        self.service_denied_timer = 60      # If deny occurred, how long did the scraper wait.
        self.wait = 7                       # Fixed waiting time for the page to load, since the page is unpredictable.
        self.completed_item = 0
        self.page_source = ''
        self.browser_closed_unexpected = False

    def restart(self):
        js_path = os.getcwd() + '\\lib\\phantomjs.exe --webdriver=4444 --load-images=no'
        agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36"
        capability = dict(DesiredCapabilities.PHANTOMJS)
        capability["phantomjs.page.settings.userAgent"] = agent
        self.js_subprocess = subprocess.Popen(
            js_path,
            stdout=subprocess.PIPE,
            creationflags=CREATE_NO_WINDOW
        )
        time.sleep(2)
        self.browser = webdriver.Remote(command_executor='http://127.0.0.1:4444/wd/hub',
                                        desired_capabilities=capability)
        self.browser.set_page_load_timeout(self.load_timeout)
        logging.info('PhantomJS launched successfully.')

    # Waiting for web page to refresh, but keep monitor if user want to abort.
    def wait_refresh(self, interval):
        for i in range(interval*2):
            if self.abort:
                raise exceptions.RuntimeError
            time.sleep(0.5)

    def get_url(self, url):
        try:
            self.browser_closed_unexpected = False
            print self.browser.window_handles
            self.browser.get(url)
            self.wait_refresh(self.wait)
            self.page_source = self.browser.page_source
        except WebDriverException:
            self.browser_closed_unexpected = True
            self.abort = True
            logging.error('Browser was accidentally closed. Abort Scraper.')
        except:
            logging.error('<get_url> method crashed: ' + traceback.format_exc())

    # self.browser sometimes will hung up, which cause the browser.get() and browser.page_source freeze the scraper.
    # Since the content tag does not include all require info, we need to parse each field.
    # Parse each required fields separately and combine into one string, write into result.txt.
    def scraper(self, search_link, source_name, idx):
        try:
            self.completed_item = idx + 1
            idx = str(idx + 1) + ','
            self.service_denied = False
            logging.info('Search Started. Target: ' + idx + source_name.encode('utf-8') + '\n')
            t_get_url = Thread(target=self.get_url, args=(search_link,), name='getURL_Thread')
            t_get_url.start()
            t_get_url.join(timeout=self.load_timeout)
            # Terminate scraper and load_url.
            if t_get_url.isAlive():
                logging.info('Timeout Waiting for browser to open the URL. Exit scraper.')
                raise common.exceptions.TimeoutException
            if self.browser_closed_unexpected:
                return
            # First check if request is identified as robot:
            logging.info('Search Results Acquired. Start Robot Check with BeautifulSoup...')
            search_page_content = self.page_source
            search_soup = BeautifulSoup(search_page_content, 'html.parser')
            search_soup_text = search_soup.text
            robot_signal = 'antirobot'
            forbidden_signal = 'Forbidden'

            if search_soup_text is None:
                self.service_denied = True
                self.service_denied_count += 1
                logging.info('No content in Search page.')
                return

            if robot_signal in search_soup_text.encode('utf-8'):
                output = idx + source_name.encode('utf-8') + u', 被识别为机器人! \n'.encode('utf-8')
                print output
                q.put(output)
                write_file(output, self.dst_path)
                logging.error(idx + source_name.encode('utf-8') +
                              ': Treated as robot. Page Source: \n' + ('#' * 150 + '\n') * 5 +
                              search_page_content.encode('utf-8') + '\n' + ('#' * 150 + '\n') * 5)
                self.service_denied = True
                self.service_denied_count += 1
                return

            if forbidden_signal in search_soup_text.encode('utf-8'):
                output = idx + source_name.encode('utf-8') + u', 403 Forbidden! \n'.encode('utf-8')
                print output
                q.put(output)
                write_file(output, self.dst_path)
                logging.error(idx + source_name.encode('utf-8') +
                              ': 403 Forbidden. Page Source: \n' + ('#' * 150 + '\n') * 5 +
                              search_page_content.encode('utf-8') + '\n' + ('#' * 150 + '\n') * 5)
                self.service_denied = True
                self.service_denied_count += 1
                return

            # Browse the first result in search page if there are any.
            if search_soup.find(class_='query_name') is not None:
                logging.info(idx + source_name.encode('utf-8') + ': Found valid entry in search page.')
                entry_link = search_soup.find(class_='query_name').get('href')
                logging.info('Try to access the Company Page link: ' + entry_link)
                t_get_url = Thread(target=self.get_url, args=(entry_link,), name='getURL_Thread')
                t_get_url.start()
                t_get_url.join(timeout=self.load_timeout)
                if t_get_url.isAlive():
                    logging.info('Timeout Waiting for browser to open the URL. Exit scraper.')
                    raise common.exceptions.TimeoutException
                if self.browser_closed_unexpected:
                    return
                logging.info('Company Page Source acquired. Start Parsing with BeautifulSoup...')
                info_page_content = self.page_source
                page_soup = BeautifulSoup(info_page_content, 'html.parser')
                # Handle the robot detection occurred after click.
                if robot_signal in page_soup.text.encode('utf-8'):
                    output = idx + source_name.encode('utf-8') + u', 被识别为机器人! \n'.encode('utf-8')
                    print output
                    q.put(output)
                    write_file(output, self.dst_path)
                    logging.error(idx + source_name.encode('utf-8') +
                                  ': Treated as robot. Page Source: \n' + ('#' * 150 + '\n') * 5 +
                                  info_page_content.encode('utf-8') + '\n' + ('#' * 150 + '\n') * 5)
                    self.service_denied = True
                    self.service_denied_count += 1
                    return
                # Extract info from the new web page.
                logging.info('Extraction started!!!')
                result = self.extract_info(page_soup)
                combined = self.info_combiner(result)
                # Record page source for investigation if not all field is extracted.
                if 'Null' in combined:
                    logging.error(idx + source_name.encode('utf-8') +
                                  ': Not all field extracted. Page Source: \n' + ('#' * 150 + '\n') * 5 +
                                  info_page_content.encode('utf-8') + '\n' + ('#' * 150 + '\n') * 5)
                logging.info('Extraction completed!!!')
                output = idx + source_name.encode('utf-8') + ',' + combined + '\n'
                write_file(output, self.dst_path)
                print output
                q.put(output)
            # Quit if no result returned. (i.e. no tag has class name = query_name)
            else:
                output = idx + source_name.encode('utf-8') + ',' + 'No result Found.\n'
                print output
                q.put(output)
                write_file(output, self.dst_path)
                logging.info(idx + source_name.encode('utf-8') + ': No valid entry in Search Page.\n')
                # logging.error(idx + source_name.encode('utf-8') +
                #               ': No valid entry in Search Page. Page Source: \n' + ('#' * 150 + '\n') * 5 +
                #               info_page_content.encode('utf-8') + '\n' + ('#' * 150 + '\n') * 5)
        # Throw by wait_refresh(), help quickly stop the scraper when user abort.
        except exceptions.RuntimeError:
            logging.info('User aborted during wait_refresh.')
            output = idx + source_name.encode('utf-8') + ',' + 'User Aborted.\n'
            print output
            write_file(output, self.dst_path)
            q.put(output)
        # Throw by self.browser.page_source,
        # happened only when killing subprocess and accessing page_source at the same time.
        except urllib2.URLError:
            logging.info('Killing subprocess and accessing page_source at the same time')
            output = idx + source_name.encode('utf-8') + ',' + 'User Aborted..\n'
            print output
            write_file(output, self.dst_path)
            q.put(output)
        # happened when webdriver took too long to respond.
        except common.exceptions.TimeoutException:
            logging.error('Selenium page_load or script execution timeout.')
            output = idx + source_name.encode('utf-8') + ',' + 'Timeout loading webpage..\n'
            print output
            write_file(output, self.dst_path)
            q.put(output)
            self.service_denied = True
            self.service_denied_count += 1
        except:
            logging.error(idx + source_name.encode('utf-8') +
                          ': Exception during webpage parsing.  \n' + traceback.format_exc() + '\n' +
                          ('#' * 150 + '\n') * 5 + self.page_source.encode('utf-8') +
                          '\n' + ('#' * 150 + '\n') * 5)
            output = idx + source_name.encode('utf-8') + ',' + 'Exception.\n'
            print output
            write_file(output, self.dst_path)
            q.put(output)
            self.service_denied = True
            self.service_denied_count += 1

    # Extract page info
    def extract_info(self, page_soup):
        result = []
        # PART 1: company name
        try:
            # Target - <div class="company_info_text"> - <p>
            company_info = page_soup.find_all('div', class_="company_info_text")
            tag = u'公司名称:'
            if len(company_info) > 0:
                temp = company_info[0].contents[0].text
                res = tag + temp
                result.append(res.encode('utf-8'))
            else:
                result.append((tag + u'未公开').encode('utf-8'))
        except:
            logging.error('Extract company_info failed\n' + traceback.format_exc())

        # PART 2: legal person name
        try:
            # Target - <td class="td-legalPersonName-value c9"> - <a>
            legal_person_name = page_soup.find_all('a', attrs={"ng-if": "company.legalPersonName", })
            tag = u'法人:'
            if len(legal_person_name) > 0:
                res = tag + legal_person_name[0].text
                result.append(res.encode('utf-8'))
            else:
                result.append((tag + u'未公开').encode('utf-8'))
        except:
            logging.error('Extract legal_person_name failed\n' + traceback.format_exc())

        # PART 3: reg capital
        try:
            # Target - <td class="td-regCapital-value"> - <p>
            reg_capital = page_soup.find_all('div', class_="baseinfo-module-content-value ng-binding")
            tag = u'注册资本:'
            if len(reg_capital) > 0:
                temp = reg_capital[0].text
                if temp == '-':
                    temp = u'未公开'
                res = tag + temp
                result.append(res.replace(',', ' ').encode('utf-8'))
            else:
                result.append((tag + 'reg_capital Null').encode('utf-8'))
        except:
            logging.error('Extract reg_capital failed\n' + traceback.format_exc())

        # PART 4: reg time
        try:
            # Target - <td class="td-regTime-value"> - <p>
            reg_time = page_soup.find_all('div', class_="baseinfo-module-content-value ng-binding")
            tag = u'注册时间:'
            if len(reg_time) > 0:
                res = tag + reg_time[1].text
                result.append(res.encode('utf-8'))
            else:
                result.append((tag + 'reg_time Null').encode('utf-8'))
        except:
            logging.error('Extract reg_time failed\n' + traceback.format_exc())

        # PART 5: staff list
        try:
            # Target - <div ng-if="company.staffList.length>0" class="ng-scope"> - <a> & <span>
            # Exclude - similar node that has "id"="nav-main-staff"
            staff_node = page_soup.find_all('div', class_="staffinfo-module-content-title")
            if len(staff_node) > 0:
                name_res = [u'任职人员:'.encode('utf-8'),]
                title_res = [u'职务:'.encode('utf-8'),]
                for name in staff_node:
                    name_res.append(name.text.encode('utf-8'))
                result.append(name_res)
                title_res.append(u'董事'.encode('utf-8'))
                result.append(title_res)
            else:
                result.append((u'任职人员:' + 'staff_list Null').encode('utf-8'))
        except:
            logging.error('Extract staff_list failed\n' + traceback.format_exc())

        # PART 6: investor list
        try:
            # Target - <div ng-if="company.investorList.length>0" class="ng-scope"> - <a>
            # Exclude - similar node that has "id"="nav-main-investment"
            investor_node = page_soup.find_all('a', attrs={"event-name": "company-detail-investment"})
            if len(investor_node) > 0 and investor_node[0] is not None:
                investor_res = [u'股东:'.encode('utf-8'),]
                for investor in investor_node:
                    investor_res.append(investor.text.encode('utf-8'))
                result.append(investor_res)
            else:
                result.append((u'股东:' + 'investor_list Null').encode('utf-8'))
        except:
            logging.error('Extract investor_list failed\n' + traceback.format_exc())
        return result

    def info_combiner(self, result):
        combined = ''
        for item in result:
            if type(item) != list:
                combined += item + ','
            else:
                combined += item[0]
                for element in item[1:]:
                    combined += element + ','
        combined = combined.strip('\n')
        return combined

    # Load the file specified by user. Generate corresponding url.
    def load_file(self):
        # Read from file, store each line (company name) into a list.
        self.format_error = False
        self.company_list = []
        bom = '%EF%BB%BF'  # Bom header for UTF-8. Need to take out before generate url.
        f = None
        try:
            f = open(self.src_path, 'r')
            for line in f:
                if bom in urllib2.quote(line):
                    line = line[3:]
                self.company_list.append(line.decode('utf-8').strip('\n'))
        # Watch out for decode error. If the input file is not encoded in UTF-8, notify user.
        except UnicodeDecodeError:
            self.format_error = True
        finally:
            f.close()

        # Compose full url list from source list.
        self.url_list = search_link_generator(self.company_list)

    # Totally 4 scenarios need to be covered.
    # 1. Scraper finished searching all items and exit normally.
    # 2. User abort when executing inside scraper(), need to wait until it finished.
    # 3. Service denied and wait for 5 min, user abort during the waiting period.
    # 4. Service denied too many times, no point to continue, the scraper exit immediately.
    # This will guarantee 'terminate()' to be called in any of the scenarios.
    def run_scraper(self):
        logging.info('Scraper officially started.')

        for i in range(len(self.url_list)):
            if not self.abort:
                self.scraper(self.url_list[i], self.company_list[i], i)
                # Restart PhantomJS every loop (to Enhance Stability)
                try:
                    self.browser.quit()
                    self.js_subprocess.terminate()
                except:
                    print "Error terminate browser or subprocess."
                    logging.info("Error terminate browser or subprocess.")
                self.restart()
                # <Scenario 4>: Service Denied too many times, terminate scraper.
                if self.service_denied_count >= self.service_denied_limits:
                    logging.info('Mission Abort! Blocked too many times!')
                    break
                # <Scenario 3>: Service Denied, wait for 1 minute to recover / user help.
                if self.service_denied:
                    # Keep monitoring the abort signal during the waiting period.
                    for cnt in range(self.service_denied_timer):
                        time.sleep(1)
                        if self.abort:
                            break
        # <Scenario 1>: Scraper finished searching all items and exit normally.
        if not self.abort:
            self.terminate()

    # Terminate the scraper thread by first set running flag to False so that new iteration will not start.
    # Then waiting for the current iteration to finish,
    # i.e. the moment closing flag become True, the current iteration is completed, so browser can be destroyed safely.
    def terminate(self):
        self.abort = True
        print 'Stopping Scraper_Thread...'
        logging.info('Stopping Scraper_Thread...')
        try:
            self.browser.quit()
            self.js_subprocess.terminate()
            print 'Scraper_Thread has stopped.'
            logging.info('Scraper_Thread has stopped.')
        except:
            print "Error terminate browser or subprocess."
            logging.info("Error terminate browser or subprocess.")


class UI(object):

    def read_file_location(self):
        self.current_status.set('Status: Loading...')
        self.status.config(fg='blue')
        self.source_path = askopenfilename(filetypes=[("Text files", "*.txt")])
        self.file_menu.entryconfig("Step #3: Post Processing", state='disable')
        if len(self.source_path) > 0:
            # Instantiate InfoScraper Class.
            # If InfoScraper already exist, that is because user already load a file, and trying to reload another one
            # We need to quit the current browser
            if self.myScraper is not None:
                self.myScraper.terminate()
            self.result_path = 'result_' + str(time.strftime('%Y %m %d %H%M', time.localtime(time.time()))) + '.txt'
            self.myScraper = InfoScraper(self.source_path, self.result_path)
            self.myScraper.load_file()
            # Bring the GUI to the front.
            self.root.lift()
            self.root.attributes('-topmost', True)
            self.root.attributes('-topmost', False)
            # Invalid file format, need to quit the current browser.
            if self.myScraper.format_error:
                self.warning()
                self.current_status.set('Status: Oops... Looks like your file is not encoded in UTF-8, '
                                        'please SAVE AS UTF-8 and retry.')
                self.status.config(fg='red')
                self.file_menu.entryconfig("Step #2: Run Forest Run!!!", state='disable')
                self.myScraper.terminate()
            else:
                self.file_menu.entryconfig("Step #2: Run Forest Run!!!", state='normal')
                self.current_status.set('Status: Source file validated. Ready to Launch. Now let\'s go to <Step #2>!')
                self.status.config(fg='blue')
        else:
            self.current_status.set('Status: No file selected.')
            self.status.config(fg='red')
            self.file_menu.entryconfig("Step #2: Run Forest Run!!!", state='disable')
            print 'No file was chosen.'

    def launch_scraper(self):
        # Clear queue every time launch the scraper.
        global q
        q = Queue.Queue()
        # Start the scraper in a new thread
        t_scraper = Thread(target=self.myScraper.run_scraper, name='Scraper_Thread')
        t_scraper.start()
        print '<Scraper_Thread> started.'
        logging.info('<Scraper_Thread> started.')

        self.current_status.set('Status: Up and Running! Yay~ The first search result will arrive very soon!')
        self.status.config(fg='red')
        self.file_menu.entryconfig("Step #2: Run Forest Run!!!", state='disable')
        self.file_menu.entryconfig("Step #1: Choose Source File", state='disable')
        self.menu_bar.entryconfig('Abort', state='normal')

        # Start Monitor the Scraper, retrieve current status in real-time, and destroy Scraper after job done.
        t_update = Thread(target=self.monitor, args=(t_scraper,), name='Monitor_Thread')
        t_update.start()
        print '<Monitor_Thread> started.'
        logging.info('<Monitor_Thread> started.')

    # A independent Thread that keeps monitoring the execution of scraper thread.
    # Terminate itself either if user abort scraper thread, or scraper thread is completed.
    def monitor(self, t_scraper):
        timer_start = time.time()
        while not self.myScraper.abort:
            try:
                # Make it unblocking, so that this thread won't stuck here forever.
                # Note that the queue could be empty the moment we abort the Scraper thread,
                # if get method is blocking, this thread may never end the current loop and exit.
                current_result = q.get(False)
                if current_result is not None:
                    current_result = self.display_formatter(current_result)
                    self.current_status.set(u'实时搜索结果: '.encode('utf-8') + current_result.replace('\n', ' ') + ' ...')
                    self.status.config(fg='black', font=("微软雅黑", 10, 'normal'))
            # When queue is empty, go to next loop immediately.
            except Queue.Empty:
                pass
            time.sleep(1)
        print 'About to exit Monitor_Thread...'
        logging.info('About to exit Monitor_Thread...')

        timer_stop = time.time()
        diff = (timer_stop - timer_start) / 60
        duration = '%.1f min. ' % diff
        print 'Total Time = ' + duration
        logging.info('Total Time = ' + duration)

        if self.myScraper.browser_closed_unexpected:
            self.current_status.set('Status: Oh no... You closed my browser, didn\'t you!' +
                                    ' (Scanned ' + str(self.myScraper.completed_item) + ' entries in ' +
                                    duration + ')')
            self.status.config(fg='red', font=("微软雅黑", 10, 'bold'))
        elif self.myScraper.service_denied_count >= self.myScraper.service_denied_limits:
            self.current_status.set('Status: Oops... Tianyancha has temporarily blocked me. ' +
                                    'Please retry if you can access Tianyancha normally.' +
                                    ' (Scanned ' + str(self.myScraper.completed_item) + ' entries in ' +
                                    duration + ')')
            self.status.config(fg='red', font=("微软雅黑", 10, 'bold'))
        else:
            self.current_status.set('Status: Job Done! Please click <Result> button to find your TXT result.' +
                                    ' Go to <Step #3> if you want!' + ' (Scanned ' + str(self.myScraper.completed_item) +
                                    ' entries in ' + duration + ')')
            self.status.config(fg='blue', font=("微软雅黑", 10, 'bold'))

        # Monitor if the scraper thread has stopped.
        for i in range(0, 10):
            print 'Scraper thread Alive = ' + str(t_scraper.isAlive())
            if t_scraper.isAlive():
                time.sleep(1)

        self.menu_bar.entryconfig('Abort', state='disable')
        self.file_menu.entryconfig("Step #3: Post Processing", state='normal')
        self.file_menu.entryconfig("Step #1: Choose Source File", state='normal')
        # When search completed or service_denied too many times, bring the GUI to top to notify user.
        self.bring_to_top()
        print 'Monitor_Thread has stopped.'
        logging.info('Monitor_Thread has stopped.')

    def display_formatter(self, res):
        # Inform delay if identified as robot or service denial.
        if '!' in res:
            res += u'自动推迟搜索60秒...如果你刚好有空的话，最好能手动登陆一下天眼查帮我识别一下验证码~'.encode('utf-8')

        # Shorten the display string if it has more than 4 colons.
        if ':' in res:
            cnt = res.count(':')
            if cnt >= 4:
                pos = 0
                # Find the 4th colon, then find the first comma after it. Extract only the string before this comma.
                for i in range(4):
                    pos = res.index(':', pos) + 1
                pos = res.index(',', pos)
                res = res[:pos]
        return res

    def on_closing(self):
        self.quit()

    def abort(self):
        self.current_status.set('Status: Cleaning up! Please wait...')
        self.status.config(fg='red')
        # If 'abort' is False, it means the scraper is still running, so we need to terminate it from outside.
        # If 'abort' is True, it means the scraper itself is already being terminated from its inside. So do nothing.
        if not self.myScraper.abort:
            self.myScraper.terminate()
        self.current_status.set('Status: Abort Successfully.')
        self.status.config(fg='black', font=("微软雅黑", 10, 'bold'))
        self.file_menu.entryconfig("Step #3: Post Processing", state='normal')
        self.file_menu.entryconfig("Step #1: Choose Source File", state='normal')

    def quit(self):
        self.current_status.set('Status: Shutting down~ Please wait...')
        self.status.config(fg='red')
        if self.myScraper is None:
            self.root.quit()
        elif self.myScraper.format_error:
            self.root.quit()
        elif not self.myScraper.abort:
            self.myScraper.terminate()
            self.root.quit()
        else:
            self.root.quit()

    def post_process(self):
        self.current_status.set('Status: Formatting the result file...')
        self.status.config(fg='red', font=("微软雅黑", 10, 'bold'))

        f = open('result\\' + self.result_path, 'r')
        self.processed_path = self.result_path.replace('result', 'processed')
        self.processed_path = self.processed_path.replace('txt', 'csv')
        marked_file = open('result\\' + self.processed_path, 'a')
        process_content = []
        for line in f:
            zhiwu = u'职务'.encode('utf-8')
            gudong = u'股东'.encode('utf-8')
            if zhiwu in line:
                start = line.find(zhiwu)
                end = line.find(gudong)
                if start != -1:
                    line = line[:start] + line[end:]
                    process_content.append(line)
            else:
                process_content.append(line)

        process_content2 = []
        for line in process_content:
            if ':' in line:
                # Extract the search result string
                org_name = line[(line.index(',') + 1):line.index(',', line.index(',') + 1)]
                search_name = line[(line.index(':') + 1):line.index(',', line.index(':'))]
                if org_name == search_name:
                    process_content2.append(self.add_mark(line, True))
                else:
                    process_content2.append(self.add_mark(line, False))
            else:
                process_content2.append(line)

        process_content3 = []
        table_header = u'序号,原始名称,@,搜索结果,@是否匹配,@,法人,@,注册资本,@,注册时间,@,任职人员,@,股东\n'
        process_content3.append(table_header.encode('utf-8'))
        for line in process_content2:
            header = []
            header.append(u'公司名称:'.encode('utf-8'))
            header.append(u'法人:'.encode('utf-8'))
            header.append(u'注册资本:'.encode('utf-8'))
            header.append(u'注册时间:'.encode('utf-8'))
            header.append(u'任职人员:'.encode('utf-8'))
            header.append(u'股东:'.encode('utf-8'))
            rep = '@,'
            removed = line
            for item in header:
                removed = removed.replace(item, rep)
            process_content3.append(removed)

        # Tricky Encoding problem: Excel will read csv in ANSI encoding, so we have to encode the UTF-8 line.
        for line in process_content3:
            marked_file.write(line.decode('utf-8').encode('mbcs'))
        f.close()
        marked_file.close()
        self.remove_duplicate(process_content3)
        self.current_status.set('Status: Post-Processing Complete!!! ' +
                                'Please Click <Result> Button to find your processed CSV file. Enjoy~')
        self.status.config(fg='blue', font=("微软雅黑", 10, 'bold'))
        self.file_menu.entryconfig("Step #3: Post Processing", state='disable')

    def remove_duplicate(self, processed_content):
        deduplicate_content = []
        header = u'序号,原始名称,@,搜索结果,@是否匹配,@,人名列表\n'
        deduplicate_content.append(header.encode('utf-8'))
        for line in processed_content[1:]:
            splitted = line.replace('\n', '').split(',@,')  # Remove next line first for better splitting.
            field_num = len(splitted)
            if field_num < 5:
                deduplicate_content.append(line)
            # Pass the first three if only five segment. (No employee or investor)
            elif field_num == 5:               
                deduplicate_content.append(',@,'.join(splitted[:3]) + '\n')
            elif field_num >= 6:
                legal_person = splitted[2]
                employees = splitted[5].split(',')
                name_list = []
                name_list.append(legal_person)
                name_list += employees
                # Add 7th segment if possible
                if field_num == 7:
                    investors = splitted[6].split(',')
                    name_list += investors
                # Clean name_list first, remove element that contain 'Null'
                name_list = [name for name in name_list if 'Null' not in name]
                if len(name_list) > 0 and name_list != ['']:
                    # Remove duplicate with one-liner
                    name_clean = {}.fromkeys(name_list).keys()
                    name_clean = filter(None, name_clean)  # Remove empty element.
                    name_clean = [x.replace(',', '') for x in name_clean]  # Remove extra comma
                    name_clean = ','.join(name_clean)
                else:
                    name_clean = "Null"
                splitted[2] = name_clean  # replace the 3rd with new name list.
                deduplicate_content.append(',@,'.join(splitted[:3]) + '\n')

        duplicate_removal_path = self.result_path.replace('result', 'processed')
        duplicate_removal_path = duplicate_removal_path.replace('.txt', '_remove_duplicate.csv')
        duplicate_remove = open('result\\' + duplicate_removal_path, 'a')
        for line in deduplicate_content:
            duplicate_remove.write(line.decode('utf-8').encode('mbcs'))
        duplicate_remove.close()

    def add_mark(self, info, marker):
        location = 1 + info.index(',', info.index(':'))
        if marker:
            marked = info[:location] + 'ok' + info[(location-1):]
        else:
            marked = info[:location] + 'False' + info[(location-1):]
        return marked

    def shortcut(self):
        folder = os.getcwd() + '\\result'
        if not os.path.exists(folder):
            os.makedirs('result')
        webbrowser.open(folder)

    def about(self):
        tkMessageBox.showinfo('About This Software', 'Any questions, please contact your husband.')

    def warning(self):
        tkMessageBox.showwarning('Input File Format Error',
                                 u'诶呀...发现了一个问题...\n\n似乎你选择的这个txt文件的编码格式不是UTF-8。\n'
                                 u'解决办法: 很简单。\n请用记事本打开你的txt文件，在菜单栏点击<文件> - <另存为>，'
                                 u'然后在弹出的新窗口的右下角有一个叫<编码>的下拉列表，选<UTF-8>保存，再尝试这个新文件即可。'
                                 u'\n\n谢谢啦...')

    def bring_to_top(self):
        self.root.focus_force()

    def __init__(self):
        self.source_path = ''
        self.result_path = ''
        self.processed_path = ''
        self.myScraper = None
        self.root = Tkinter.Tk()
        self.root.title('Awesome Scraper ' + __version__)
        ico_path = os.getcwd() + '\\ico\\icon.ico'
        if os.path.exists(ico_path):
            self.root.iconbitmap(ico_path)
        self.root.geometry("1000x65")
        # Create menu_bar in root.
        self.menu_bar = Tkinter.Menu(self.root)

        # Create sub menu 'file_menu' of menu_bar.
        self.file_menu = Tkinter.Menu(self.menu_bar, tearoff=0)
        # Add command to the sub menu.
        self.file_menu.add_command(label="Step #1: Choose Source File",
                                   command=self.read_file_location)
        self.file_menu.add_command(label="Step #2: Run Forest Run!!!",
                                   command=self.launch_scraper, state='disabled')
        self.file_menu.add_command(label="Step #3: Post Processing",
                                   command=self.post_process, state='disabled')
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.quit)
        # Add sub menu 'Start' into menu_bar.
        self.menu_bar.add_cascade(label="Start", menu=self.file_menu, font=10)
        # Add sub menu 'Abort' into menu_bar.
        self.menu_bar.add_command(label='Abort', command=self.abort, state='disabled')
        # Add sub menu 'Path' into menu_bar.
        self.menu_bar.add_command(label='Result', command=self.shortcut, state='normal')
        # Create sub menu 'help_menu' of menu_bar.
        self.help_menu = Tkinter.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="About", command=self.about, font=("微软雅黑", 10))
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)

        self.root.config(menu=self.menu_bar)

        self.current_status = Tkinter.StringVar()
        self.current_status.set('Status: Ready. Let\'s go to Start Menu - <Step #1>.')

        self.status = Tkinter.Label(self.root,
                                    textvariable=self.current_status,
                                    bd=20,
                                    font=("微软雅黑", 10, 'bold'),
                                    fg="blue",
                                    width=960,
                                    height=10,
                                    wraplength=960,
                                    anchor='w',
                                    justify='left')
        self.status.pack(side='left')

        # Register a Protocol, so that when click the X button, it will first terminate all thread.
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

if not os.path.exists('result'):
    os.makedirs('result')
logging.basicConfig(filename='result\\error.log',
                    filemode='w',
                    format='[%(asctime)s] - [%(levelname)s] >>> %(message)s',
                    level=logging.INFO)
q = Queue.Queue()
myUI = UI()
