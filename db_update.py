from pymongo import MongoClient
from datetime import date
from bson.objectid import ObjectId
import logging



class DBAgent:
        def __init__(self, host="localhost", port=27017, dbname='icair'):
                self.dbhost = host
                self.dbport = port
                self.dbname = dbname
                self.client = MongoClient(self.dbhost, self.dbport)
                self.db = self.client[self.dbname]

        def __del__(self):
                self.client.close()

        def db_insert(self, coll_name, new_obj):
                # add current date
                new_obj["current_date"] = datetime.datetime.utcnow()

                collection = self.db[coll_name]
                logging.debug("Inserting object into collection: %s", coll_name)
                return collection.insert(new_obj)


        def db_find(self, coll_name, query={}):
                collection = self.db[coll_name]
                posts = collection.find(query)
                count = posts.count()
                if (count == 0):
                        logging.info("No matching records found in the collection: %s", coll_name)
                        return None
                elif (count == 1):
                        logging.info("Found 1 matching record in the collection: %s", coll_name)
                        return posts[0]

                else:
                        logging.info("Found %d matching records in the collection: %s", count, coll_name)
                        objs = []
						for post in posts:
                                objs.append(post)
                        return objs

logging.basicConfig()
dba = DBAgent()
print dba.db_find('events', {"_id":ObjectId('53c5429f414e7376f260ef25')})
