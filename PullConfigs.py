from Database import ExcelFunctions
from netmiko import BaseConnection
import os


class PullConfigs:
    divide = "========================="

    def __init__(self, doc=ExcelFunctions.__new__(ExcelFunctions), ssh=BaseConnection.__new__(BaseConnection)):
        self.ssh = ssh
        self.doc = doc
        self.host = self.doc.current_ip
        # file_name = f'{self.doc.current_ip} - {ssh.find_prompt().split("#")[0]}.txt'
        self.data = ""
        self.divide = self.divide

    @staticmethod
    def open_file(directory='', file_name='', mode="w+"):
        if os.path.exists(directory):
            os.chdir(directory)  # Change working directory
        else:
            print(f"ERROR: Directory not found: {directory}\nBut I'll make it for you anyway :)")
            os.mkdir(directory)

        os.chdir(f'{directory}')
        file = open(f'{file_name}', mode)
        return file

    def get_configs(self):
        ssh = self.ssh
        vendor = self.doc.get_cell('Vendor', self.doc.host_info[self.host]["cell"]).value
        cmd_directory = f'{self.doc.directory}\Command Sets'
        cmd_file_name = f'{vendor}.txt'
        cmd_file = self.open_file(cmd_directory, cmd_file_name, 'r')

        data = ""
        cmds = cmd_file.read().split("\n")

        for cmd in cmds:
            data = data + f'{self.divide} {cmd} START {self.divide}\n'
            data = data + ssh.send_command(cmd)
            data = data + f'\n{self.divide} {cmd} END {self.divide}\n'

        cmd_file.close()

        host_directory = f'{self.doc.directory}\Configs'
        host_file_name = f'{self.doc.current_ip} - {ssh.find_prompt().split("#")[0]}.txt'
        host_file = self.open_file(host_directory, host_file_name, 'w+')

        host_file.write(data)
        host_file.close()
