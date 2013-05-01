#!/usr/bin/env python

import sys, urllib2

# try:
#   from argparse import ArgumentParser
# except ImportError:
url = "http://argparse.googlecode.com/files/argparse-1.2.1.tar.gz"
file_name = url.split('/')[-1]
u = urllib2.urlopen(url)
f = open(file_name, 'wb')
meta = u.info()
file_size = int(meta.getheaders("Content-Length")[0])
print "Downloading: %s Bytes: %s" % (file_name, file_size)

file_size_dl = 0
block_sz = 8192
while True:
    buffer = u.read(block_sz)
    if not buffer:
        break

    file_size_dl += len(buffer)
    f.write(buffer)
    status = r"%10d  [%3.2f%%]" % (file_size_dl, file_size_dl * 100. / file_size)
    status = status + chr(8)*(len(status)+1)
    print status,

f.close()
# url = "http://argparse.googlecode.com/files/argparse-1.2.1.tar.gz"
# urllib.urlretrieve(url, "argparse-1.2.1.gz")
  # pass
