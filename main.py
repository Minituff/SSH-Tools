from datetime import datetime
from SSHConnection import SSH
from STIGs import Vuln
from Database import ExcelFunctions
from PullConfigs import PullConfigs

startTime = datetime.now()


def login():
    #Database.main()

    xlDoc = ExcelFunctions()

    for host in xlDoc.host_info.keys():
        ip = xlDoc.host_info[host]["ip"]
        vendor = xlDoc.host_info[host]["vendor"]
        run = xlDoc.host_info[host]["run"]

        xlDoc.current_ip = host
        #ssh = SSHConection.SSH(host, xlDoc)
        ssh = SSH(doc=xlDoc)

        #print(type(ssh.sess))
        if ssh.sess is not None:  #If it could log in
            pull_configs = PullConfigs(xlDoc, ssh.sess)
            pull_configs.get_configs()

    stigs = Vuln(xlDoc)
    xlDoc.save_workbook()


login()


endTime = datetime.now()
totalTime = endTime - startTime
print(f'\nScript finished successfully!\n'
      f'Run Time: {totalTime}')
