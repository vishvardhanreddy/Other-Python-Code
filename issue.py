#!/usr/bin/python

import re
import json

class issue:
        def __init__(self):
                self.log_sections = ["log"]

        def get_log_sections(self):
                return self.log_sections;

        def parse_log(self, section_data):
                regex = "\s*(\S+\s+\S+\s+\S+\s+\S+)\s+(\S+)\s+(.*)"
                m = re.match(regex, section_data)
                if (m):
                        return {"log":{"timestamp":m.group(1), "process":m.group(2), "message":m.group(3)}}

        def parse_version(self, section_data):
                print

        def parse_install(self, section_data):
                print

class issue1(issue):
        def __init__(self):
                self.log_sections = ["before", "log", "after", "version", "install"]

        def parse_show_proc(self, section_data):
                header=None
                values=None
                my_log_obj = {}

                for line in section_data.split("\n"):
                        if (line.find("JID") != -1):
                                header=line
                        elif (line.find("sysdb_mc") != -1):
                                values=line

                        if (header and values ):
                                header_list = re.split("\s+", header.strip())
                                values_list = re.split("\s+", values.strip())
 for i in (range(len(header_list))):
                                        if (header_list[i] == "Dynamic"):
                                                my_log_obj["heap_size"] = values_list[i]
                                        elif (header_list[i] == "JID"):
                                                my_log_obj["JID"] = values_list[i]

                return my_log_obj

        def parse_before(self, section_data):
                my_obj = self.parse_show_proc(section_data)
                return {"before" : my_obj}

        def parse_after(self, section_data):
                my_obj = self.parse_show_proc(section_data)
                return {"after": my_obj}


i = issue1()

print i.get_log_sections()

print i.parse_log_entry("2014 Apr 3 10:41:20.194 sysdb_mc[416]: %SYSDB-SYSDB-3-MEMORY : malloc of 34329876 bytes failed")

print i.parse_before('''Fri May 16 17:55:21.196 EST
JID          Text       Data      Stack    Dynamic Process
------ ---------- ---------- ---------- ---------- -------
416        188416      65536      81920    4288512 sysdb_mc''')

print i.parse_after('''Fri May 16 17:55:21.196 EST
JID          Text       Data      Stack    Dynamic Process
------ ---------- ---------- ---------- ---------- -------
416        188416      65536      81920    4288512 sysdb_mc''')
                                                        