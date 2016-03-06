# -*- coding: utf-8 -*-
# Author: Michael Yuan
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import urllib2
import random
from threading import Thread
from Tkinter import *
from tkFileDialog import askopenfilename
from tkMessageBox import *
import Queue
import exceptions


# 关于线程设计的一个关键特性就是避免在里面出现任何blocking的等待特性功能，
# 因为一旦线程中有了blocking的等待，这个线程就失去了在任何时间都能实时响应外部的请求的能力，
# 这时候就会出现各种bug，例如线程无法接收外部信号量来关闭自己。
# 如果线程中必须有等待，那么一定要把它做成non-blocking的轮询机制。
# 例如sleep N秒做成每sleep 1秒查一次是否被关闭，Queue.get也要配成non-blocking，然后每秒去查并立即返回。

# Read from file, return company name as list.
def read_input_list(filename):
    company_name = []
    bom = '%EF%BB%BF'
    try:
        f = open(filename, 'r')
        for line in f:
            # print line
            if bom in urllib2.quote(line):
                line = line[3:]
            company_name.append(line.decode('utf-8').strip('\n'))
    finally:
        f.close()
    return company_name


# Generate UTF-8 encoded url link of the search path.
def search_link_generator(entrylist):
    link_list = []
    domain = 'http://www.tianyancha.com/search/'
    for entry in entrylist:
        info = entry.encode('utf-8')
        link = domain + urllib2.quote(info)
        link_list.append(link)
        # print link
    return link_list


def write_file(info, path='result.txt'):
    f = open(path, 'a')
    try:
        f.write(info)
    finally:
        f.close()


def rand_wait(low, up):
    return random.randint(low, up)


# Extract page info
def extract_info(page_soup):
    result = []
    # company name
    try:
        company_info = page_soup.find_all('div', class_="row b-c-white show_in_pc")
        if len(company_info) > 0:
            res = u'公司名称:' + company_info[0].p.string
            result.append(res.encode('utf-8'))
        else:
            result.append('company_info Null')
    except:
        write_file('Extract company_info failed', 'extract_error.log')

    # legal person name
    try:
        legal_person_name = page_soup.find_all('td', class_="td-legalPersonName-value c9")
        if len(legal_person_name) > 0:
            res = u'法人:' + legal_person_name[0].a.string
            result.append(res.encode('utf-8'))
        else:
            result.append('legal_person_name Null')
    except:
        write_file('Extract legal_person_name failed', 'extract_error.log')

    # reg capital
    try:
        reg_capital = page_soup.find_all('td', class_="td-regCapital-value")
        if len(reg_capital) > 0:
            res = u'注册资本:' + reg_capital[0].p.string
            result.append(res.strip(',').encode('utf-8'))
        else:
            result.append('reg_capital Null')
    except:
        write_file('Extract reg_capital failed', 'extract_error.log')

    # reg time
    try:
        reg_time = page_soup.find_all('td', class_="td-regTime-value")
        if len(reg_time) > 0:
            res = u'注册时间:' + reg_time[0].p.string
            result.append(res.encode('utf-8'))
        else:
            result.append('reg_time Null')
    except:
        write_file('Extract reg_time failed', 'extract_error.log')

    # staff list
    try:
        staff_node = page_soup.find_all('div', class_="row b-c-white", style="padding-left:2px;")
        staff_soup = BeautifulSoup(str(staff_node[0]), 'html.parser')
        staff_name = staff_soup.find_all('a')
        staff_title = staff_soup.find_all('span')
        name_res = [u'任职人员:'.encode('utf-8'),]
        title_res = [u'职务:'.encode('utf-8'),]
        if len(staff_name) > 0:
            for name in staff_name:
                name_res.append(name.string.encode('utf-8'))
        else:
            name_res.append('staff_name Null')
        result.append(name_res)
        if len(staff_title) > 0:
            for title in staff_title:
                title_res.append(title.string.strip('\n').encode('utf-8'))
        else:
            title_res.append('staff_title Null')
        result.append(title_res)
    except:
        write_file('Extract staff_list failed', 'extract_error.log')

    # investor list
    try:
        investor_node = page_soup.find_all('div', class_="row b-c-white")
        investor_soup = BeautifulSoup(str(investor_node[3]), 'html.parser')
        investor_name = investor_soup.find_all('a')
        investor_res = [u'股东:'.encode('utf-8'),]
        exception = u'案件'.encode('utf-8')
        if len(investor_name) > 0:
            for investor in investor_name:
                info = investor.string.encode('utf-8')
                if exception not in info:
                    investor_res.append(info)
                else:
                    investor_res = [u'股东:'.encode('utf-8'), 'investor_title Null']
                    break
        else:
            investor_res.append('investor_title Null')

        result.append(investor_res)
    except:
        write_file('Extract investor_list failed', 'extract_error.log')

    return result


def info_combiner(result):
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


class InfoScraper(object):
    def __init__(self, src_path, dst_path):
        self.src_path = src_path
        self.dst_path = dst_path
        self.browser = webdriver.PhantomJS()
        self.wait = 5
        self.timer_start = 0
        self.timer_stop = 0
        self.duration = ''
        self.abort = False
        self.closing = False
        self.service_denied = False
        self.service_denied_count = 0
        self.completed_item = 0

    # Waiting for web page to refresh, but keep monitor if user want to abort.
    def wait_refresh(self, interval):
        for i in range(interval):
            if self.abort:
                raise exceptions.RuntimeError
            time.sleep(1)

    # Since the content tag does not include all require info, we need to parse each field.
    # Parse each required fields separately and combine into one string, write into result.txt.
    def scraper(self, search_link, source_name, idx):
        try:
            self.completed_item = idx + 1
            idx = str(idx) + ','
            self.service_denied = False
            self.browser.get(search_link)
            self.wait_refresh(self.wait)
            # Handle scenario when request is identified as robot:
            search_soup = BeautifulSoup(self.browser.page_source, 'html.parser')
            if search_soup.find('div', class_='center') is not None:
                output = idx + source_name.encode('utf-8') + \
                         ', Identified as Robot! Current search aborted, wait for 5min.\n'
                print output
                q.put(output)
                write_file(output, self.dst_path)
                write_file(output + self.browser.page_source.encode('utf-8'), 'error.log')
                self.service_denied = True
                self.service_denied_count += 1
                return
            if search_soup.find('title') is not None and '403' in search_soup.find('title').string:
                output = idx + source_name.encode('utf-8') + \
                         ', 403 Forbidden! Current search aborted, wait for 5min.\n'
                print output
                q.put(output)
                write_file(output, self.dst_path)
                write_file(output + self.browser.page_source.encode('utf-8'), 'error.log')
                self.service_denied = True
                self.service_denied_count += 1
                return
            # Click the first result if there are result returned.
            if search_soup.find(class_='query_name') is not None:
                self.browser.find_element_by_class_name('query_name').click()
                self.wait_refresh(self.wait)
                # Extract info from the new web page.
                page_soup = BeautifulSoup(self.browser.page_source, 'html.parser')
                result = extract_info(page_soup)
                combined = info_combiner(result)
                output = idx + source_name.encode('utf-8') + ',' + combined + '\n'
                write_file(output, self.dst_path)
                q.put(output)
                print output
            # Quit if no result returned. (i.e. no tag has class name = query_name)
            else:
                output = idx + source_name.encode('utf-8') + ',' + 'No result Found.\n'
                print output
                q.put(output)
                write_file(output, self.dst_path)
                write_file(output + self.browser.page_source.encode('utf-8'), 'error.log')
        except exceptions.RuntimeError:
            output = idx + source_name.encode('utf-8') + ',' + 'User Aborted.\n'
            write_file(output, self.dst_path)
            q.put(output)
        except:
            output = idx + source_name.encode('utf-8') + ',' + 'Exception.\n'
            write_file(output, self.dst_path)
            q.put(output)
            write_file(output + self.browser.page_source.encode('utf-8'), 'error.log')
        # browser.close()

    # Totally 4 scenarios need to be covered.
    # 1. Scraper finished searching all items and exit normally.
    # 2. User abort when executing inside scraper(), need to wait until it finished.
    # 3. Service denied and wait for 5 min, user abort during the waiting period.
    # 4. Service denied too many times, no point to continue, the scraper exit immediately.
    # This will guarantee 'terminate()' to be called in any of the scenarios.
    def start_scraper(self):
        self.timer_start = time.time()
        # Read item from source list.
        source_list = read_input_list(self.src_path)
        # Compose full url list from source list.
        url_list = search_link_generator(source_list)
        # Use selenium to launch each url_list.
        for i in range(len(source_list)):
            if not self.abort:
                self.scraper(url_list[i], source_list[i], i)
                # <Scenario 3>: Service Denied, wait for 5 minute to recover.
                if self.service_denied:
                    # Keep monitoring the abort signal during the waiting period.
                    for cnt in range(300):
                        time.sleep(1)
                        if self.abort:
                            break
                # <Scenario 4>: Service Denied too many time, no point to continue, the scraper exit immediately.
                if self.service_denied_count > 5:
                    break
            # <Scenario 2>: User abort when executing inside scraper(), need to wait until it finished.
            if self.abort:
                self.closing = True  # Detecting current iteration is finished and ready for cleanup.
                break

        self.timer_stop = time.time()
        diff = (self.timer_stop - self.timer_start) / 60
        self.duration = '%.1f min. ' % diff
        print 'Total Time = ' + self.duration

        # <Scenario 1>: Scraper finished searching all items and exit normally.
        if not self.closing:
            self.closing = True
            self.terminate()

    # Terminate the scraper thread by first set running flag to False so that new iteration will not start.
    # Then waiting for the current iteration to finish,
    # i.e. the moment closing flag become True, the current iteration is completed, so browser can be destroyed safely.
    def terminate(self):
        self.abort = True
        print '<Scraper_Thread> terminate start.'
        for i in range(1, 20):  # Keep checking if scraper is about to stop the loop.
            if self.closing:
                self.browser.close()
                print '<Scraper_Thread> terminate completed.'
                break
            else:
                time.sleep(0.5)


class UI(object):

    def read_file_location(self):
        self.source_path = askopenfilename(filetypes=[("Text files", "*.txt")])
        self.file_menu.entryconfig("Step #3: Post Processing", state='disable')
        if len(self.source_path) > 0:
            self.file_menu.entryconfig("Step #2: Run Forest Run!!!", state='normal')
            self.current_status.set('Status: Source file selected. Ready to Launch. Now let\'s go to <Step #2>!')
            self.status.config(fg='blue')
        else:
            self.current_status.set('Status: No file selected.')
            self.status.config(fg='red')
            print 'No file was chosen.'

    def launch_scraper(self):
        # Instantiate InfoScraper Class.
        self.result_path = 'result_' + str(time.strftime('%Y %m %d %H%M', time.localtime(time.time()))) + '.txt'
        self.myScraper = InfoScraper(self.source_path, self.result_path)
        self.myScraper.interval = 5
        # Start the scraper in a new thread
        t = Thread(target=self.myScraper.start_scraper, name='Scraper_Thread')
        t.start()

        self.current_status.set('Status: Up and Running! Yay~ The first search result will arrive very soon!')
        self.status.config(fg='red')
        self.file_menu.entryconfig("Step #2: Run Forest Run!!!", state='disable')
        self.menu_bar.entryconfig('Abort', state='normal')

        # Start Monitor the Scraper, retrieve current status in real-time, and destroy Scraper after job done.
        t_update = Thread(target=self.monitor, name='UI_Monitor_Thread')
        t_update.start()

    # A independent Thread that keeps monitoring the execution of scraper thread.
    # Terminate itself either if user abort scraper thread, or scraper thread is completed.
    def monitor(self):
        while not self.myScraper.abort:
            try:
                # Make it unblocking, so that this thread won't stuck here forever.
                # Note that the queue could be empty the moment we abort the Scraper thread,
                # if get method is blocking, this thread may never end the current loop and exit.
                current_result = q.get(False)
                if current_result is not None:
                    current_result = self.display_cutter(current_result)
                    self.current_status.set('Status: ' + current_result.strip('\n') + ' ......')
                    self.status.config(fg='black', font=("微软雅黑", 10, 'normal'))
            # When queue is empty, go to next loop immediately.
            except Queue.Empty:
                pass
            time.sleep(1)
        print '<UI_Monitor_Thread> terminate completed.'
        if self.myScraper.service_denied_count > 5:
            self.current_status.set('Status: Oops... Looks like Tianyancha has temporarily blocked me. Help me!!!' +
                                    ' (Scanned ' + str(self.myScraper.completed_item) + ' entries in ' +
                                    self.myScraper.duration + ')')
            self.status.config(fg='red', font=("微软雅黑", 10, 'bold'))
        else:
            self.current_status.set('Status: Job Done! Please find the <result.txt> under the same folder.' +
                                    ' Go to <Step #3> if you want!' + ' (Scanned ' + str(self.myScraper.completed_item) +
                                    ' entries in ' + self.myScraper.duration + ')')
            self.status.config(fg='blue', font=("微软雅黑", 10, 'bold'))
        print '<UI_Monitor_Thread> Scraper Destroyed.'
        self.myScraper = None
        self.menu_bar.entryconfig('Abort', state='disable')
        self.file_menu.entryconfig("Step #3: Post Processing", state='normal')

    def display_cutter(self, res):
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

    def quit(self):
        self.current_status.set('Status: Shutting down~ Please wait...')
        self.status.config(fg='red')
        if self.myScraper is None:
            self.root.quit()
        elif not self.myScraper.abort:
            self.myScraper.terminate()
            self.root.quit()
        else:
            self.root.quit()

    def post_process(self):
        self.current_status.set('Status: Formatting the result file...')
        self.status.config(fg='red', font=("微软雅黑", 10, 'bold'))
        f = open(self.result_path, 'r')
        self.proc_path = self.result_path.replace('result', 'processed')
        self.proc_path = self.proc_path.replace('txt', 'csv')
        marked_file = open(self.proc_path, 'a')
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
        self.current_status.set('Status: Post-Processing Complete!!! ' +
                                'Please find the <processed.csv> in the same folder. Enjoy~')
        self.status.config(fg='blue', font=("微软雅黑", 10, 'bold'))

    def add_mark(self, info, marker):
        location = 1 + info.index(',', info.index(':'))
        if marker:
            marked = info[:location] + 'ok' + info[(location-1):]
        else:
            marked = info[:location] + 'False' + info[(location-1):]
        return marked

    def about(self):
        showinfo('About This Software', 'Any questions, please contact your husband.')

    def not_implemented(self):
        showwarning('Sorry~', 'Under Construction...')

    def __init__(self):
        self.source_path = ''
        self.result_path = ''
        self.proc_path = ''
        self.myScraper = None
        self.root = Tk()
        self.root.title('Awesome Scraper v1.0')
        self.root.iconbitmap('icon.ico')
        self.root.geometry("1000x65")
        # Create menu_bar in root.
        self.menu_bar = Menu(self.root)

        # Create sub menu 'file_menu' of menu_bar.
        self.file_menu = Menu(self.menu_bar, tearoff=0)
        # Add command to the sub menu.
        self.file_menu.add_command(label="Step #1: Choose Source File",
                                   command=self.read_file_location)
        self.file_menu.add_command(label="Step #2: Run Forest Run!!!",
                                   command=self.launch_scraper, state='disabled')
        self.file_menu.add_command(label="Step #3: Post Processing",
                                   command=self.post_process, state='disabled')
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.quit)
        # Add sub menu 'file' into menu_bar.
        self.menu_bar.add_cascade(label="Start", menu=self.file_menu, font=10)

        self.menu_bar.add_command(label='Abort', command=self.abort, state='disabled')

        # Create sub menu 'help_menu' of menu_bar.
        self.help_menu = Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="About", command=self.about, font=("微软雅黑", 10))
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)

        self.root.config(menu=self.menu_bar)

        self.current_status = StringVar()
        self.current_status.set('Status: Ready.')

        self.status = Label(self.root,
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

q = Queue.Queue()
myUI = UI()

