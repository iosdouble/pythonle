import datetime
import os
import subprocess
import time

import psutil as psutil


def test2():
    fnull = open(os.devnull, 'w')
    return1 = subprocess.call('ping 127.0.0.1 -n 1', shell=True, stdout=fnull, stderr=fnull)
    rt = (return1 == 0)
    if rt:
        print(28)
    else:
        print(30)
    return rt

def sh(command):           #命令行获取硬盘信息
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    result = p.stdout.read().decode('gbk')
    return result
def sh_receive():
    disk_c = ''
    usedg_c = 0
    hardDisk_c = {}
    try:               #把 try提到for循环下。
        if os.name == 'nt':
            disks = sh("wmic logicaldisk get deviceid").split('\r\r\n')[1:-2]
            for disk in disks:
                if disk.strip():
                    res = sh('fsutil volume diskfree %s' % disk).split('\r\n')[:-2]
                    free = int(res[0].split(':')[1])
                    all_bit = int(res[1].split(':')[1])
                    total = round(all_bit / 1024 / 1024, 2)  # mb
                    totalg = int(total / 1024)  # gb
                    used = round((all_bit - free) / 1024 / 1024, 2)  # mb
                    usedg = int(used / 1024)  # gb
                    Idle = round(free / 1024 / 1024, 2)  # mb
                    ram = round(used / total, 2)
                    disk_c += "{}GB/{}GB#".format(usedg, totalg)  # 新加的
                    usedg_c += Idle
                    hardDisk_c[disk.strip()] = {"total": total, "used": used, "Idle": Idle, "ram": ram}
        return (disk_c, usedg_c, hardDisk_c)
    except Exception as e:
        print(e)
        return (disk_c, usedg_c, hardDisk_c)

def netInfo():                          #ip网卡信息
    new_ip = {}
    new_mac = {}
    cache_ip = ""
    cache_mac = ""
    dic = psutil.net_if_addrs()
    for adapter in dic:
        snicList = dic[adapter]
        mac = False  # '无 mac 地址'
        ipv4 = False  # '无 ipv4 地址'
        ipv6 = '无 ipv6 地址'
        for snic in snicList:
            if snic.family.name in {'AF_LINK', 'AF_PACKET'}:
                mac = snic.address
            elif snic.family.name == 'AF_INET':
                ipv4 = snic.address
        if ipv4 and ipv4 != '127.0.0.1':
            new_ip[adapter] = str(ipv4)
            cache_ip += '{}#'.format(str(ipv4))
        if mac:
            if len(mac) == 17:
                new_mac[adapter] = str(mac)
                cache_mac += '{}#'.format(str(mac))
    return (new_ip, new_mac, cache_ip, cache_mac)
def netSpeed():
    net2 = psutil.net_io_counters()                            #pernic=True 可以查看每一个网卡的传输量
    old_bytes_sent = round(net2.bytes_recv / 1024, 2)     #kb
    old_bytes_rcvd = round(net2.bytes_sent / 1024, 2)
    time.sleep(1)
    net3 = psutil.net_io_counters()
    bytes_sent = round(net3.bytes_recv / 1024, 2)
    bytes_rcvd = round(net3.bytes_sent / 1024, 2)
    netSpeedOutgoing = bytes_sent - old_bytes_sent
    netSpeedInComing = bytes_rcvd - old_bytes_rcvd
    return (netSpeedOutgoing, netSpeedInComing)

class ServerPicker:
    def __init__(self):
        # 网络连接
        self.onlineStatus = int(test2())
        # CPU 占用率
        self.cpuUtilization = psutil.cpu_percent(0)
        # 总内存，内存使用率，使用内存，剩余内存
        mem = psutil.virtual_memory()
        # 所有单位均为byte字节,1 KB = 1024 bytes,1 MB = 1024 KB,1 GB = 1024 MB
        # total：内存总大小   percent：已用内存百分比(浮点数)     used：已用内存      free：可用内存
        self.totalRAM = round(mem.total / 1024 / 1024, 2)  # mb
        self.totalRAMg = round(mem.total / 1024 / 1024 / 1024, 2)  # gb
        self.ramUtilization = round(mem.percent, 2)
        self.usedRAM = round(mem.used / 1024 / 1024, 2)  # mb
        self.usedRAMg = round(mem.used / 1024 / 1024 / 1024, 2)  # gb
        self.IdleRAM = round(mem.free / 1024 / 1024, 2)  # mb
        self.IdleRAMg = round(mem.free / 1024 / 1024 / 1024, 2)  # gb
        # 硬盘=个数和每个硬盘的总容量、使用容量、和剩余容量及使用率
        self.disk, self.usedg, self.hardDisk = sh_receive()
        # 系统开机时间
        # powerOnTime=datetime.datetime.fromtimestamp(psutil.boot_time()).strftime("%Y-%m-%d %H:%M:%S")
        t = int(psutil.boot_time())
        self.powerOnTime = str(datetime.datetime.fromtimestamp(t))
        # 系统开机时长
        self.powerOnDuration = int(time.time())
        # # 获取多网卡MAC地址
        # # 获取ip地址
        self.ip, self.mac, self.cache_ip, self.cache_mac = netInfo()
        # 网络收发速率
        self.netSpeedOutgoing, self.netSpeedInComing = netSpeed()

if __name__ == '__main__':
    netInfo()