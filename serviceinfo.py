import os
'''
os.statvfs方法用于返回包含文件描述符fd的文件的文件系统的信息。
语法：os.statvfs([path])
返回值
f_bsize: 文件系统块大小
f_frsize: 分栈大小
f_blocks: 文件系统数据块总数
f_bfree: 可用块数
f_bavail:非超级用户可获取的块数
f_files: 文件结点总数
f_ffree: 可用文件结点数
f_favail: 非超级用户的可用文件结点数
f_fsid: 文件系统标识 ID
f_flag: 挂载标记
f_namemax: 最大文件长度
'''
def disk_stat():
    hd={}
    disk = os.statvfs('/')
    hd['available'] = float(disk.f_bsize * disk.f_bavail)
    hd['capacity'] = float(disk.f_bsize * disk.f_blocks)
    hd['used'] = float((disk.f_blocks - disk.f_bfree) * disk.f_frsize)
    res = {}
    res['used'] = round(hd['used'] / (1024 * 1024 * 1024), 2)
    res['capacity'] = round(hd['capacity'] / (1024 * 1024 * 1024), 2)
    res['available'] = res['capacity'] - res['used']
    res['percent'] = int(round(float(res['used']) / res['capacity'] * 100))
    return res