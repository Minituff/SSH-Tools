import netmiko
from datetime import datetime
from netmiko.ssh_exception import NetMikoTimeoutException, SSHException
from Database import ExcelFunctions
#import Database


class SSH:
    def __init__(self, doc=ExcelFunctions()):
        self.doc = doc
        self.sess_device_type = None
        self.sess_username = "admin"
        self.sess_password = "cisco"
        self.sess_secret = self.sess_password
        self.sess_ip = doc.current_ip
        self.host = doc.current_ip
        self.ses_hostname = ""
        self.sess = self.ssh_connect()

    def ssh_connect(self):
        scan_status_cell = self.doc.get_cell('Scan Status', self.doc.host_info[self.host]["cell"])
        self.sess_device_type = self.doc.get_cell('Vendor', self.doc.host_info[self.host]["cell"]).value

        try:
            ssh = netmiko.ConnectHandler(device_type=self.sess_device_type,
                                         ip=self.sess_ip,
                                         username=self.sess_username,
                                         secret=self.sess_secret,
                                         password=self.sess_password,
                                         timeout=5.0,
                                         fast_cli=True,
                                         verbose=True)

            print(f'\nSuccessful connection made to {self.sess_ip} aka {ssh.find_prompt().split("#")[0]}\n')
            scan_status_cell.value = f'Connected on: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            return ssh

        except (EOFError, SSHException, NetMikoTimeoutException, ConnectionError):
            print(f'SSH connection timed out on: {self.sess_ip}')
            scan_status_cell.value = f'Failed to connect on: {datetime.now().strftime("%Y-%m-%d %H:%M")}'

    @staticmethod
    def pull_hostname(raw_name, device_type):
        pass

if __name__ == '__main__':
    pass
    #print('SSH Connection is being run by itself')
    #SSH = SSH()
else:
    pass
    #print('SSH Connection is being imported from another module')
    #SSH = SSH()