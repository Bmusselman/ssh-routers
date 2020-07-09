#! python 3
import paramiko, pyperclip, openpyxl, datetime, os

##READ ME -- CISCO LIMITATION
##3.8.3.6 -m: read a remote command or script from a file
##
##The -m option performs a similar function to the ‘Remote command’ box in the SSH
##panel of the PuTTY configuration box (see section 4.18.1).
##
##However, the -m option expects to be given a local file name, and it will read
##a command from that file.
##
##With some servers (particularly Unix systems), you can even put multiple
##lines in this file and execute more than one command in sequence, or a
##whole shell script; but this is arguably an abuse, and cannot be expected to
##work on all servers. In particular, it is known not to work with
##certain ‘embedded’ servers, such as Cisco routers.

##TODO:
##User enter BR and gives output


# time format (YYYY-MM-DD_HOUR-MIN)
timeFormat = datetime.datetime.now().strftime('%Y-%m-%d__%H-%M')

# time of program start
beginTime = datetime.datetime.now()

os.chdir(r'\\rem\cd\ITHelpDesk\Branch_Networking\Track Logs')
logFile = open(timeFormat + '.txt', 'a+')
logFile.write('Below is a list of all branch routers that are\nunerachable or showing 1+ track statements down')

# need this stuff to establish connection
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# excel doc path
excelSheet = r'\\rem\cd\ITHelpDesk\Branch_Networking\BranchIPs.xlsx'

# load excel doc
workbook = openpyxl.load_workbook(excelSheet)

# load excel sheet
sheet = workbook['Sheet1']

# authenticate - enter user/pass
# need to find way to store this not in plain-text
user = ''
pwd = ''


# filler
filler = 25 * '-'

# list for sorting
track_list = []

# track statement down
track_down = 'Down'.encode() # encode to byte since output is in bytes

command_list = ['sh track']#, 'sh ip int br']#, 'sh int s0/0/0', 'sh service- s0/0/0']


def track_status(ip_address):
    output = ''
    for command in command_list:
        try:
            ssh.connect(hostname = ip_address, username = user, password = pwd, port = 22)
            stdin, stdout, stderr = ssh.exec_command(command)
            output = stdout.read()
            if track_down in output:
                logFile.write('\n' + '\n' + str(filler) + str(brNum) + str(filler))
                logFile.write(output.decode())
        except TimeoutError:
                logFile.write('\n' + '\n' + str(filler) + str(brNum) + str(filler))
                logFile.write('\n\nUnable to SSH to this branch')


for row in range(2, sheet.max_row): # for each row, starting row 2 to last
    brIP = sheet.cell(row = row, column = 2).value
    brNum = sheet.cell(row = row, column = 1).value
    if(brIP is None): 
        continue
    track_status(brIP)


# print program execution time
execTime = datetime.datetime.now() - beginTime
print('The script execution time is: ' + str(execTime))
logFile.write('\n\nThe script execution time is: ' + str(execTime))
logFile.close()
    
