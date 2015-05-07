import logging

from pymongo import MongoClient

class DBAgent:
        def __init__(self, host="localhost", port=27017, dbname='icair'):
                self.dbhost = host
                self.dbport = port
                self.dbname = dbname
                client = MongoClient(self.dbhost, self.dbport)
                self.db = client[self.dbname]



        def db_find(self, coll_name, query={}):
                collection = self.db[coll_name]
                posts = collection.find(query)
                count = posts.count()
                if (count == 0):
                        print "no records found"
                        return None
                elif (count == 1):
                        print "found 1 record"
                        return posts[0]

                else:
                        print "found more records"
                        logging.info("Found %d matching records in the collection: %s", count, coll_name)
                        objs = []
                        for post in posts:
                                objs.append(post)
                        return objs



log_file = "test.log"
log_level = logging.DEBUG

logging.basicConfig(filename=log_file, level=log_level, format='[%(asctime)s] [%(levelname)s] %(module)s::%(funcName)s: %(message)s', datefmt='%a, %b %d %Y %H:%M:%S')
logging.info('Started monitoring agent')


dba = DBAgent()
objs = dba.db_find('events',{"issue_id":"cscuo10444"})

print objs

