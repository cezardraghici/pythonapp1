import sqlite3
import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename

def connect():
    conn=sqlite3.connect("material.db")
    cur=conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS materiale (id INTEGER, nume text, firma_produs text, unitate text, nr_bucati INTEGER, pret INTEGER)")
    conn.commit()
    conn.close()

def drop():
   conn=sqlite3.connect("material.db")
   cur=conn.cursor()
   cur.execute("DROP TABLE materiale")
   conn.commit()
   conn.close()

def insert_from_excel():
    conn=sqlite3.connect("material.db")
    filepath = askopenfilename(
        filetypes=[("Excels Files", "*.xlsx"), ("All Files", "*.*")]
    )
    wb = pd.read_excel(filepath,sheet_name = 0)
    wb.to_sql("materiale",conn,if_exists='append', index=False)
    conn.commit()
    conn.close()


def insert(nume,firma_produs,unitate,nr_bucati,pret):
    conn=sqlite3.connect("material.db")
    cur=conn.cursor()
    cur.execute("INSERT INTO materiale VALUES(NULL,?,?,?,?,?)",(nume,firma_produs,unitate,nr_bucati,pret))
    conn.commit()
    conn.close()


def delete():
    conn=sqlite3.connect("material.db")
    cur=conn.cursor()
    cur.execute("DELETE FROM materiale")
    conn.commit()
    conn.close()

def view():
    conn=sqlite3.connect("material.db")
    cur=conn.cursor()
    cur.execute("SELECT nume,firma_produs,unitate,nr_bucati,pret FROM materiale group by nume")
    rows=cur.fetchall()
    conn.close()
    return rows

def search(nume="", firma_produs=""):
    conn=sqlite3.connect("material.db")
    cur=conn.cursor()
    cur.execute("SELECT distinct nume,firma_produs,unitate,nr_bucati,max(pret) FROM materiale where nume like ? and firma_produs like ? group by nume order by pret desc",('%'+nume+'%','%'+firma_produs+'%'),)
    rows=cur.fetchall()
    conn.close()
    return rows

# def idprod(nume,pret):
#     conn=sqlite3.connect("material.db")
#     cur=conn.cursor()
#     cur.execute("SELECT id FROM materiale where nume = ? AND pret = ? and data=?",(nume, pret))
#     rows=cur.fetchall()
#     conn.close()
#     return rows

#drop()
connect()





