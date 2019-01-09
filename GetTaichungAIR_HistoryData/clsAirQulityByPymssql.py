import pymssql # http://pymssql.org/en/latest/index.html 
from Setting import Config as conf

class AirQulity():

    @staticmethod   # python @classmethod and @ staticmethod difference, http://missions5.blogspot.tw/2014/12/python-classmethod-and-staticmethod.html
    def SaveValue(dicAirQulity,SiteEName,PublishDateTime):
        
        ## get DB connection information, http://pymssql.org/en/latest/pymssql_examples.html
        Server = conf.Value('DB','ServerIP')            
        User = conf.Value('DB','User')
        Password = conf.Value('DB','Password')   
        DBName = conf.Value('DB','DBName')  

        try:
            mConn = pymssql.connect(Server, User, Password, DBName)

            #print(dicAirQulity)

            cursor = mConn.cursor()
            dictionary_with_update = {
                'param1': dicAirQulity.get("AQI"),
                'param2': dicAirQulity.get("O3"),
                'param3': dicAirQulity.get("PM25"),
                'param4': dicAirQulity.get("PM10"),
                'param5': dicAirQulity.get("CO"),
                'param6': dicAirQulity.get("SO2"),
                'param7': dicAirQulity.get("NO2"),
                'param8': PublishDateTime,
                'param9': SiteEName,

            }

            cursor = mConn.cursor()
            cursor.execute('UPDATE tbAirSite SET \
                            AQI  = %(param1)s, \
                            O3   = %(param2)s, \
                            PM25 = %(param3)s, \
                            PM10 = %(param4)s, \
                            CO   = %(param5)s, \
                            SO2  = %(param6)s, \
                            NO2  = %(param7)s,  \
                            PublishTime  = %(param8)s  \
                            WHERE SiteEName = %(param9)s',dictionary_with_update)
            mConn.commit()
            result = 1

        except:
            print("AirQulity::SaveValue error...")
            result = 0

        finally:
            cursor.close()
            mConn.close()

        return result

