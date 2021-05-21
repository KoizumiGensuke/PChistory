import sys, win32com.client
import sqlite3

def GetDetails(folder):
    """
    Get date, name, url from entries in target folder including sub-folders
    
    2021/5/21 koizumi
    """
    result = []
    for item in folder.Items():
        if item.IsFolder:
            ret = GetDetails(item.GetFolder) # recurcive call
            if ret is not None:
                result += ret
        else:
            url = folder.GetDetailsOf(item, 0)
            name = folder.GetDetailsOf(item, 1)
            date = folder.GetDetailsOf(item, 2)
            result.append([date, name, url])
    return(result)

shell = win32com.client.Dispatch('Shell.Application')
folder = shell.Namespace(0x22)
ret = GetDetails(folder)

cn = sqlite3.connect('PChistory.db')
cur = cn.cursor()

cur.execute("""CREATE TABLE IF NOT EXISTS history (
    date TIMESTAMP NOT NULL,
    name TEXT,
    url TEXT);""")
cn.commit()

# use temporary table to remove records which exist already in db.
cur.execute("""CREATE TEMPORARY TABLE IF NOT EXISTS temp (
    date TIMESTAMP NOT NULL,
    name TEXT,
    url TEXT);""")
cn.commit()

for i in range(len(ret)):
    cur.execute("""INSERT INTO temp (date, name, url) VALUES (?, ?, ?);""", 
                [ret[i][0], ret[i][1], ret[i][2]])
    cn.commit()

# add new records
cur.execute("""INSERT INTO history 
    SELECT * FROM temp WHERE NOT EXISTS (
        SELECT date FROM history 
        WHERE temp.date=history.date AND temp.url=history.url);""")
cn.commit()

cur.close()
cn.close()
