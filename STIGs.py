from Database import ExcelFunctions
from netmiko import BaseConnection
from openpyxl.comments import Comment
from openpyxl.styles import Alignment
from PullConfigs import PullConfigs
import os
import re


class Vuln:
    def __init__(self, doc=ExcelFunctions.__new__(ExcelFunctions)):

        self.doc = doc
        #self.host = self.doc.current_ip
        # file_name = f'{self.doc.current_ip} - {ssh.find_prompt().split("#")[0]}.txt'
        self.data = ""
        self.file = None
        self.current_ip = ""
        self.hostname = ""
        self.divide = PullConfigs.divide

        directory = f'{self.doc.directory}\Configs'
        self.cycle_files(directory, "r")

    def run_stigs(self):
        self.pull_hostname()
        self.pull_version()

    def cycle_files(self, directory='', mode="w+"):
        if os.path.exists(directory):
            os.chdir(directory)  # Change working directory
        else:
            print(f"ERROR: Directory not found: {directory}\nBut I'll make it for you anyway :)")
            os.mkdir(directory)

        os.chdir(f'{directory}')

        for filename in os.listdir(directory):
            if filename.endswith('.txt'):
                with open(os.path.join(directory, filename), mode) as file:
                    self.current_ip = filename.split(" -")[0]
                    self.hostname = filename.split("- ")[1].split(".txt")[0]
                    self.file = file
                    self.run_stigs()
                    file.close()

    def re_match(self, file, cmd):
        data = file.read()
        # match = re.search(r'(^START[\s\S]+^END)', data, re.M | re.S)
        search = r'(^' + re.escape(f'{self.divide} {cmd} START {self.divide}') + '[\s\S]+^' + re.escape(
            f'{self.divide} {cmd} END {self.divide}') + ')'

        match = re.search(search, data, re.M | re.S)
        if match:
            data = match.group(1)
            data = data.split(f'{self.divide} {cmd} START {self.divide}')[1]
            data = data.split(f'{self.divide} {cmd} END {self.divide}')[0]

        return data

    @staticmethod
    def re_include(data, include_list):
        # Functions like the "pipe to include" for Cisco.
        # A group of data is inputted, and a list of data is matched against it
        # If multiple objects are added to the list, it is matched as a logical OR.

        match_data = ""
        for char in include_list:
            if match_data is "":
                match_data = ".*" + char
            else:
                match_data = match_data + ".*|" + char
        match_data = match_data + ".*"

        result = ""
        for line in data.split('\n'):
            match = re.search(match_data, line, flags=re.IGNORECASE)
            if match:
                result = result + "\n" + match.group(0)

        return result

    def pull_version(self):
        cell = self.doc.get_cell('Version', self.doc.host_info[self.current_ip]["cell"])
        cell.value = None
        cmd = "show version"

        output_raw = self.re_match(self.file, cmd)
        output = self.re_include(output_raw, ["version"])

        output = output.split(" ")
        version = ""

        for each_word in output:
            if "version" in each_word.lower():
                version = output[output.index(each_word) + 1]
                break

        version = version.strip(",")
        print(version)
        cell.value = version
        cell_cmt = Comment(f'{self.hostname}# {cmd}\n{output_raw}', '')
        cell_cmt.width = 1000
        cell_cmt.height = 500
        cell.comment = cell_cmt

    def pull_hostname(self):
        cell = self.doc.get_cell('Hostname', self.doc.host_info[self.current_ip]["cell"])
        cell.value = self.hostname

    def pull_serial(self):
        cell = self.doc.get_cell('Serial Number', self.doc.host_info[self.current_ip]["cell"])
        cell.alignment = Alignment(wrap_text=True)
        cell.value = None

        cmd = "show version | i [Ss]ystem [Ss]erial [Nn]umber"

        output_raw = self.re_match(self.file, cmd)

        test = cmd + "\n" \
                     "System Serial Number               : FCW2035D0CY\n" \
                     "System Serial Number               : FOC2035U0GC"

        output = re.split('\s+', test)

        for each_word in output:
            if each_word is ":":
                if cell.value is None:
                    cell.value = output[output.index(each_word) + 1]
                else:
                    cell.value = f'{cell.value}\n{output[output.index(each_word) + 1]}'

        cell_cmt = Comment(f'{self.hostname}# {cmd}\n{output_raw}', '')
        cell_cmt.width = 1000
        cell_cmt.height = 500
        cell.comment = cell_cmt
