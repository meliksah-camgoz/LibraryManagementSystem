# -*- coding: utf-8 -*-
import sys
import sqlite3
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import pandas as pd
import openpyxl
import os
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt 
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import *
from AnaSayfaUI import *
from HakkindaUI import *
from girissayfasi import * 
from kayitol import * 
from sayfalarUI import *
from istatistiklerUI import *

QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)

app=QtWidgets.QApplication(sys.argv)
MainWindow=QtWidgets.QMainWindow()
ui= Ui_MainWindow() 
ui.setupUi(MainWindow)

app2=QtWidgets.QApplication(sys.argv)
MainWindow2=QtWidgets.QMainWindow()
ui2= Ui_hakkinda() 
ui2.setupUi(MainWindow2)


#giriş sayfası için bağlantılar
app3=QtWidgets.QApplication(sys.argv)
MainWindow3=QtWidgets.QMainWindow()
ui3= Ui_girispenceresi() 
ui3.setupUi(MainWindow3)

MainWindow3.show()

#kayıt sayfası için bağlantılar
app4=QtWidgets.QApplication(sys.argv)
MainWindow4=QtWidgets.QMainWindow()
ui4= Ui_kayitoll() 
ui4.setupUi(MainWindow4)

app5=QtWidgets.QApplication(sys.argv)
MainWindow5=QtWidgets.QMainWindow()
ui5= Ui_sayfalar() 
ui5.setupUi(MainWindow5)

app6=QtWidgets.QApplication(sys.argv)
MainWindow6=QtWidgets.QMainWindow()
ui6= Ui_istatistik()
ui6.setupUi(MainWindow6)

import sqlite3
global curs
global conn
conn=sqlite3.connect('veritabani.db')
curs=conn.cursor()
sorguCreTblkitaplik=("CREATE TABLE IF NOT EXISTS kitap(                         \
                     Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,                      \
                     Nu TEXT NOT NULL UNIQUE,                         \
                     Adi TEXT NOT NULL,                        \
                     YazarAdi TEXT NOT NULL,                             \
                     Yayinevi TEXT NOT NULL,                             \
                     Tarih TEXT NOT NULL,                            \
                     Tur TEXT NOT NULL,                         \
                     Raf TEXT NOT NULL,                              \
                     Ayrac TEXT NOT NULL,                             \
                     Dil TEXT NOT NULL,                            \
                     Icerik TEXT                                     )")

curs.execute(sorguCreTblkitaplik)
conn.commit()

conn=sqlite3.connect('veritabani.db') 
curs=conn.cursor() 
sorguEmanetlerr=("CREATE TABLE IF NOT EXISTS emanetlerr(                         \
                     Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,                      \
                     Birlik TEXT NOT NULL,                         \
                     Rütbe TEXT NOT NULL,                        \
                     AdSoyad TEXT NOT NULL,                             \
                     TelNo TEXT NOT NULL,                             \
                     KitapNo TEXT NOT NULL,                            \
                     KitapAdi TEXT NOT NULL,                         \
                     TeslimTarihi TEXT NOT NULL,                              \
                     IadeTarihi TEXT NOT NULL                                    )")
curs.execute(sorguEmanetlerr)
conn.commit()

#giriş-kayıt veritabanı bağlantısı
conn2=sqlite3.connect('GirisVeri.db') 
curs2=conn2.cursor() 
curs2.execute("CREATE TABLE IF NOT EXISTS GirisVeri (`ıd2`	INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,username	TEXT NOT NULL UNIQUE,passwoord	TEXT NOT NULL,adsoyad	TEXT NOT NULL)")
conn2.commit() #bağlantıyı commit ediyoruz.


def girisiac():
    MainWindow.close()
    MainWindow2.close()
    MainWindow3.show()
    MainWindow4.close()
    MainWindow5.close()
   
def kayitiac():
    MainWindow.close()
    MainWindow2.close()
    MainWindow4.show()
    MainWindow3.close()
    MainWindow5.close()
  
def hakkindaac():
    MainWindow2.show()

def istatistikAc():
    istatistikListele()
    MainWindow6.show()
    MainWindow5.close()    

def kayitol():
    
    ad=ui4.lineEditAdSoyad.text()
    kullanici=ui4.lineEditKulAdi.text()
    sifre1=ui4.lineEditSifre.text()
    sifre2=ui4.lineEditSifreTekrar.text()
    

    if (sifre1==sifre2)and((sifre1!="")and(sifre2!="")and(kullanici!="")) :
        curs2.execute('INSERT INTO GirisVeri (username, passwoord, adsoyad ) VALUES ( ?, ?, ?)', (kullanici,sifre1,ad) )
        conn2.commit()
        ui3.statusbar.showMessage(" "*5 + " Kayıt gerçekleştirildi.", 1500)
        ui4.lineEditAdSoyad.clear() 
        ui4.lineEditKulAdi.clear()
        ui4.lineEditSifre.clear()
        ui4.lineEditSifreTekrar.clear()
        girisiac()
    else:
        ui4.statusbar.showMessage(" "*5 + " Şifreler Farklı. ", 1500)
              
def giris():
   
    global username
    username=ui3.lineEdit_2.text()
    passwoord=ui3.lineEdit.text()
    
    curs2.execute("SELECT * FROM GirisVeri WHERE username='%s' and passwoord='%s'" %(username,passwoord))
    conn2.commit()
    
    if(len(curs2.fetchall())>0):  
        ui.statusbar.showMessage(" "*3 + " Giriş yapıldı", 1500) 
        MainWindow3.close() 
        MainWindow5.show()  
        ui3.lineEdit_2.clear()  
        ui3.lineEdit.clear()
      
        
    else:
        ui3.statusbar.showMessage(" "*3 + "Hatalı Kullanıcı Adı/Şifre", 1500)  

def anasayfa():
    MainWindow5.close()
    MainWindow.show()

def geriIstatistik():
    MainWindow6.close()
    MainWindow5.show()    
def geriMenu():
    MainWindow.close()
    MainWindow5.show()
def sifre_goster():

    if ui3.checkBox.isChecked():
        ui3.lineEdit.setEchoMode(QLineEdit.Normal)    
    else:
        ui3.lineEdit.setEchoMode(QLineEdit.Password)             

def tabloya_satir_ekle():
        kitap_no = ui.lineEdit_kitapNo.text()
        kitap_adi = ui.lineEdit_kitapAdi.text()
        teslim_tarihi = ui.dateEditTeslim.text()
        iade_tarihi = ui.dateEditIade.text()
        ad_soyad = ui.lineEdit_adSoyad.text()
        tel_no = ui.lineEdit_telNo.text()
        rutbe = ui.lineEdit_rutbe.text()
        birlik = ui.cmbEdit_birlik.currentText()
        
        curs.execute("INSERT INTO emanetlerr \
                         (Birlik,Rütbe,AdSoyad,TelNo,KitapNo,KitapAdi,TeslimTarihi,IadeTarihi) \
                          VALUES(?,?,?,?,?,?,?,?)",\
                              (birlik,rutbe,ad_soyad,tel_no,kitap_no,kitap_adi,teslim_tarihi,iade_tarihi,))
    
        conn.commit()
        LISTELEe()

def EKLE():
    _lneNu=ui.lneNu.text()
    _lneAdi=ui.lneAdi.text()
    _lneYazarAdi=ui.lneYazarAdi.text()
    _lneYayinevi=ui.lneYayinevi.text()
    _cmbTarih=ui.cmbTarih.currentText()
    _lneTur=ui.lneTur.text()
    _spnRaf=ui.spnRaf.value()
    _cmbAyrac=ui.cmbAyrac.currentText()
    _lneDil=ui.lneDil.text()
    _lneIcerik=ui.lneIcerik.toPlainText()
            
    curs.execute("INSERT INTO kitap \
                         (Nu,Adi,YazarAdi,Yayinevi,Tarih,Tur,Raf,Ayrac,Dil,Icerik) \
                          VALUES(?,?,?,?,?,?,?,?,?,?)",\
                              (_lneNu,_lneAdi,_lneYazarAdi,_lneYayinevi,_cmbTarih,_lneTur,_spnRaf,_cmbAyrac,_lneDil,_lneIcerik,))
    
    conn.commit()
    LISTELE()
    
def LISTELEe():
    
    ui.emanetler.clear()
    ui.emanetler.setHorizontalHeaderLabels(('No','Birlik','Rütbe','AdSoyad','TelNo','KitapNo', \
                                                  'KitapAdi','TeslimTarih','IadeTarih'))
    
    ui.emanetler.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    #ui.emanetler.setColumnHidden(0, True)
    curs.execute("SELECT * FROM emanetlerr")
    
    for satirIndeks, satirVeri in enumerate(curs):
        for sutunIndeks, sutunVeri in enumerate (satirVeri):
            ui.emanetler.setItem(satirIndeks,sutunIndeks,QTableWidgetItem(str(sutunVeri)))

    ui.cmbEdit_birlik.setCurrentIndex(-1)
    ui.lineEdit_rutbe.clear()
    ui.lineEdit_adSoyad.clear()
    ui.lineEdit_telNo.clear()
    ui.lineEdit_kitapAdi.clear()
    ui.lineEdit_kitapNo.clear()
    ui.dateEditIade.clear()
    ui.dateEditTeslim.clear()

LISTELEe()

def LISTELE():
    
    ui.tblwBilgi.clear()
    ui.tblwBilgi.setHorizontalHeaderLabels(('No','Numara','Kitap Ad','Yazar','Yayınevi', \
                                                  'Tarih','Tür','Raf', 'Raf Ayraç', \
                                                  'Dil', 'İçerik'))
    
    ui.tblwBilgi.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    #ui.tblwBilgi.setColumnHidden(0, True)
    curs.execute("SELECT * FROM kitap")
    
    for satirIndeks, satirVeri in enumerate(curs):
        for sutunIndeks, sutunVeri in enumerate (satirVeri):
            ui.tblwBilgi.setItem(satirIndeks,sutunIndeks,QTableWidgetItem(str(sutunVeri)))        
    
    ui.lneNu.clear()
    ui.lneAdi.clear()
    ui.lneYazarAdi.clear()
    ui.lneYayinevi.clear()
    ui.cmbTarih.setCurrentIndex(-1)
    ui.lneTur.clear()
    ui.spnRaf.setValue(55)
    ui.cmbAyrac.setCurrentIndex(-1)
    ui.lneDil.clear()
    ui.lneIcerik.clear()   

LISTELE() 

def istatistikListele():
    curs.execute("SELECT COUNT(*) FROM kitap")
    kayitSayisi=curs.fetchone()
    ui6.lblKayitSayisi.setText(str(kayitSayisi[0])) 

    curs.execute("SELECT COUNT(*) FROM emanetlerr WHERE birlik = ?", ("2'nci Tank Tabur K.lığı",))
    ikiTankTaburSayisi=curs.fetchone()
    ui6.lbl2TankTaburu.setText(str(ikiTankTaburSayisi[0]))

    curs.execute("SELECT COUNT(*) FROM emanetlerr WHERE birlik = ?", ("1'inci Tank Tabur K.lığı",))
    birTankTaburSayisi=curs.fetchone()
    ui6.lbl1TankTaburu.setText(str(birTankTaburSayisi[0]))

    curs.execute("SELECT COUNT(*) FROM emanetlerr WHERE birlik = ?", ("Lojistik Destek Tabur K.lığı",))
    lojistikTaburSayisi=curs.fetchone()
    ui6.lblLojistikDestek.setText(str(lojistikTaburSayisi[0]))

    curs.execute("SELECT COUNT(*) FROM emanetlerr WHERE birlik = ?", ("Komuta Destek Tabur K.lığı",))
    KomutaDestekTaburSayisi=curs.fetchone()
    ui6.lblKomutaDestek.setText(str(KomutaDestekTaburSayisi[0]))

    curs.execute("SELECT COUNT(*) FROM emanetlerr WHERE birlik = ?", ("Topçu Tabur K.lığı",))
    TopcuTaburSayisi=curs.fetchone()
    ui6.lblTopcu.setText(str(TopcuTaburSayisi[0]))

    curs.execute("SELECT COUNT(*) FROM emanetlerr WHERE birlik = ?", ("Mekanize Piyade Tabur K.lığı",))
    MekanizeTaburSayisi=curs.fetchone()
    ui6.lblMekanize.setText(str(MekanizeTaburSayisi[0]))     

def CIKIS():
    cevap=QMessageBox.question(MainWindow,"ÇIKIŞ","Programdan çıkmak istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        conn.close()       
        sys.exit(app.exec_())
    else:
        MainWindow.show()

def CIKIS2():
    cevap=QMessageBox.question(MainWindow,"ÇIKIŞ","Programdan çıkmak istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        conn.close()       
        sys.exit(app.exec_())
    else:
        MainWindow5.show()             


def SIL():
    cevap=QMessageBox.question(MainWindow,"KAYIT SİL","Kaydı silmek istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        secili=ui.tblwBilgi.selectedItems()
        silinecek=secili[1].text()
        try:
            curs.execute("DELETE FROM kitap WHERE Nu='%s'" %(silinecek))
            conn.commit()
            
            LISTELE()
            
            ui.statusbar.showMessage("KAYIT SİLME İŞLEMİ BAŞARIYLA GERÇEKLEŞTİ...",10000)
        except Exception as Hata:
            ui.statusbar.showMessage("Şöyle bir hata ile karşılaşıldı:"+str(Hata))
    else:
        ui.statusbar.showMessage("Silme işlemi iptal edildi...",10000)

def IADE():
    cevap=QMessageBox.question(MainWindow,"KAYIT SİL","Kitabı İade Almak İstediğinize Eminmisiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        secili=ui.emanetler.selectedItems()
        silinecek=secili[5].text()
        try:
            curs.execute("DELETE FROM emanetlerr WHERE KitapNo='%s'" %(silinecek))
            conn.commit()
            
            LISTELEe()
            
            ui.statusbar.showMessage("BAŞARIYLA KİTAP İADE ALINDI...",10000)
        except Exception as Hata:
            ui.statusbar.showMessage("Şöyle bir hata ile karşılaşıldı:"+str(Hata))
    else:
        ui.statusbar.showMessage("Silme işlemi iptal edildi...",10000)        
               
def KITAPARA():
    adi=ui.lneAra.text()
    curs.execute("SELECT * FROM kitap WHERE LOWER (Adi)='%s'" %(adi))
    conn.commit()
    ui.tblwBilgi.clear()
    for row,columnvalues in enumerate(curs):
        for column, value in enumerate(columnvalues):
            ui.tblwBilgi.setItem(row, column, QTableWidgetItem(str(value)))
    ui.tblwBilgi.setHorizontalHeaderLabels(("No",'Numara','Kitap Adı','Yazar', \
                                                  'Yayınevi', 'Yayın Tarihi', 'Tür',  \
                                                  'Raf No', 'Raf Ayraç','Dil','İçerik'))
    ui.tblwBilgi.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)        
        
def YAZARARA():
    YazarAdi=ui.lneAra.text()
    curs.execute("SELECT * FROM kitap WHERE LOWER (YazarAdi)='%s'" %(YazarAdi))
    conn.commit()
    ui.tblwBilgi.clear()
    for row,columnvalues in enumerate(curs):
        for column, value in enumerate(columnvalues):
            ui.tblwBilgi.setItem(row, column, QTableWidgetItem(str(value)))
    ui.tblwBilgi.setHorizontalHeaderLabels(("No",'Numara','Kitap Adı','Yazar', \
                                                  'Yayınevi', 'Yayın Tarihi', 'Tür',  \
                                                  'Raf No', 'Raf Ayraç','Dil','İçerik'))
    ui.tblwBilgi.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)       

def NUMARA():
    numara=ui.lneAra.text()
    curs.execute("SELECT * FROM kitap WHERE Nu='%s'" %(numara))
    conn.commit()
    ui.tblwBilgi.clear()
    for row,columnvalues in enumerate(curs):
        for column, value in enumerate(columnvalues):
            ui.tblwBilgi.setItem(row, column, QTableWidgetItem(str(value)))
    ui.tblwBilgi.setHorizontalHeaderLabels(("No",'Numara','Kitap Adı','Yazar', \
                                                  'Yayınevi', 'Yayın Tarihi', 'Tür',  \
                                                  'Raf No', 'Raf Ayraç','Dil','İçerik'))
    ui.tblwBilgi.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)        

def TURARA():
    tur=ui.lneAra.text()
    curs.execute("SELECT * FROM kitap WHERE LOWER (Tur)='%s'" %(tur))
    conn.commit()
    ui.tblwBilgi.clear()
    for row,columnvalues in enumerate(curs):
        for column, value in enumerate(columnvalues):
            ui.tblwBilgi.setItem(row, column, QTableWidgetItem(str(value)))
    ui.tblwBilgi.setHorizontalHeaderLabels(("No",'Numara','Kitap Adı','Yazar', \
                                                  'Yayınevi', 'Yayın Tarihi', 'Tür',  \
                                                  'Raf No', 'Raf Ayraç','Dil','İçerik'))
    ui.tblwBilgi.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)        

def DILARA():
    dil=ui.lneAra.text()
    curs.execute("SELECT * FROM kitap WHERE LOWER (Dil)='%s'" %(dil))
    conn.commit()
    ui.tblwBilgi.clear()
    for row,columnvalues in enumerate(curs):
        for column, value in enumerate(columnvalues):
            ui.tblwBilgi.setItem(row, column, QTableWidgetItem(str(value)))
    ui.tblwBilgi.setHorizontalHeaderLabels(("No",'Numara','Kitap Adı','Yazar', \
                                                  'Yayınevi', 'Yayın Tarihi', 'Tür',  \
                                                  'Raf No', 'Raf Ayraç','Dil','İçerik'))
    ui.tblwBilgi.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)        
                 
def DOLDUR():
    secili = ui.tblwBilgi.selectedItems()
    
    if len(secili) < 11:
        return  # Eksik seçim varsa hiçbir şey yapma
    
    ui.lneNu.setText(secili[1].text())
    ui.lneAdi.setText(secili[2].text())
    ui.lneYazarAdi.setText(secili[3].text())
    ui.lneYayinevi.setText(secili[4].text())
    ui.cmbTarih.setCurrentText(secili[5].text())
    ui.lneTur.setText(secili[6].text())
    ui.spnRaf.setValue(int(secili[7].text()))
    ui.cmbAyrac.setCurrentText(secili[8].text())
    ui.lneDil.setText(secili[9].text())
    ui.lneIcerik.setText(secili[10].text())

def DOLDUR2():
    secili = ui.emanetler.selectedItems()
    
    if len(secili) < 9:
        return  # Eksik seçim varsa hiçbir şey yapma
    
    ui.cmbEdit_birlik.setCurrentText(secili[1].text())
    ui.lineEdit_rutbe.setText(secili[2].text())
    ui.lineEdit_adSoyad.setText(secili[3].text())
    ui.lineEdit_telNo.setText(secili[4].text())
    ui.lineEdit_kitapNo.setText(secili[5].text())
    ui.lineEdit_kitapAdi.setText(secili[6].text())
    ui.dateEditTeslim.setText(secili[7].text())
    ui.dateEditIade.setText(secili[8].text())
       
def GUNCELLE():
    cevap=QMessageBox.question(MainWindow,"KAYIT GÜNCELLE","Kaydi güncellemek istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        try:
            
            secili=ui.tblwBilgi.selectedItems()
            _Id=int(secili[0].text())
            _lneNu=ui.lneNu.text()
            _lneAdi=ui.lneAdi.text()
            _lneYazarAdi=ui.lneYazarAdi.text()
            _lneYayinevi=ui.lneYayinevi.text()
            _cmbTarih=ui.cmbTarih.currentText()
            _lneTur=ui.lneTur.text()
            _spnRaf=ui.spnRaf.value()
            _cmbAyrac=ui.cmbAyrac.currentText()
            _lneDil=ui.lneDil.text()
            _lneIcerik=ui.lneIcerik.toPlainText()
        
                       
            curs.execute("UPDATE kitap SET Nu=?, Adi=?, YazarAdi=?, Yayinevi=?, Tarih=?, Tur=?, Raf=?, Ayrac=?, Dil=?, Icerik=? WHERE Id=?", \
                         (_lneNu,_lneAdi,_lneYazarAdi,_lneYayinevi,_cmbTarih,_lneTur,_spnRaf,_cmbAyrac,_lneDil,_lneIcerik,_Id))
            conn.commit()
            
            LISTELE()
            
        except Exception as Hata:
            ui.statusbar.showMessage("Şöyle bir hata meydana geldi"+str(Hata))
    else:
        ui.statusbar.showMessage("Güncellme iptal edildi",10000)

def GUNCELLEemanet():
    cevap=QMessageBox.question(MainWindow,"KAYIT GÜNCELLE","Kaydi güncellemek istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        try:
            
            secili=ui.emanetler.selectedItems()
            _Id=int(secili[0].text())
            kitap_no = ui.lineEdit_kitapNo.text()
            kitap_adi = ui.lineEdit_kitapAdi.text()
            teslim_tarihi = ui.dateEditTeslim.text()
            iade_tarihi = ui.dateEditIade.text()
            ad_soyad = ui.lineEdit_adSoyad.text()
            tel_no = ui.lineEdit_telNo.text()
            rutbe = ui.lineEdit_rutbe.text()
            birlik = ui.cmbEdit_birlik.currentText()
        
                       
            curs.execute("UPDATE emanetlerr SET Birlik=?, Rütbe=?, AdSoyad=?, TelNo=?, KitapNo=?, KitapAdi=?, TeslimTarihi=?, IadeTarihi=? WHERE Id=?", \
                         (birlik,rutbe,ad_soyad,tel_no,kitap_no,kitap_adi,teslim_tarihi,iade_tarihi,_Id))
            conn.commit()
            
            LISTELEe()
            
        except Exception as Hata:
            ui.statusbar.showMessage("Şöyle bir hata meydana geldi"+str(Hata))
    else:
        ui.statusbar.showMessage("Güncellme iptal edildi",10000)        
            
def HAKKINDA():
    MainWindow2.show()
     
def ANACIKIS():
    cevap=QMessageBox.question(MainWindow,"ÇIKIŞ","Kütüphane sisteminden ayrılmak istediğinize emin misiniz?", \
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        conn.close()
        sys.exit(app.exec_())
    else:
        MainWindow3.show()

import pandas as pd
import os

def EXCEL_AKTAR():
    try:
        curs.execute("SELECT Nu, Adi, YazarAdi, Yayinevi, Tarih, Tur, Raf, Ayrac, Dil, Icerik FROM kitap")
        veriler = curs.fetchall()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Kitaplar"

        basliklar = ["Sıra No", 'Numara','Kitap Adı','Yazar', 'Yayınevi', 'Yayın Tarihi', 'Tür','Raf No', 'Raf Ayraç','Dil','İçerik']
        ws.append(basliklar)

        for sira, satir in enumerate(veriler, start=1):
         ws.append([sira] + list(satir))

        masaustu = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        dosya_yolu = os.path.join(masaustu, "kitaplar.xlsx")
        wb.save(dosya_yolu)
        QMessageBox.information(MainWindow, "Başarılı", f"Excel dosyası kaydedildi:\n{dosya_yolu}")
        print(f"Excel çıktısı kaydedildi: {dosya_yolu}")

    except Exception as e:
        QMessageBox.warning(MainWindow, "Hata", f"Excel aktarma sırasında hata oluştu:\n(NOT: Exel Açıkken Kaydetme Yapmayın)\n{str(e)}")

def TUTANAK(secilen_id):
    try:
        # Veritabanına bağlan
        conn = sqlite3.connect("veritabani.db")
        curs = conn.cursor()

        secili=ui.emanetler.selectedItems()
        _Id=int(secili[0].text())

        # Seçilen ID'ye göre bilgileri al
        curs.execute("SELECT Birlik, Rütbe, AdSoyad, TelNo, KitapNo, KitapAdi, TeslimTarihi, IadeTarihi FROM emanetlerr WHERE Id = ?", (_Id,))
        veri = curs.fetchone()

        if veri is None:
            QMessageBox.warning(None, "Hata", "Seçilen kayıt bulunamadı.")
            return

        birlik, rutbe, adsoyad, telno, kitapno, kitapadi, teslimtarihi, iadetarihi = veri

        # Word belgesi oluştur
        doc = Document()
        

    # Başlık: ortalanmış ve büyük harf
        title = doc.add_paragraph("TAŞINIR MAL TESLİM TESELLÜM BELGESİ")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.runs[0]
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        run.font.size = Pt(17)
        run.font.bold = True

    # Boş satır
        doc.add_paragraph("")

    # Ana metin (tek paragraf, süslü parantez içindekiler değişken)
        text = (
         f"{teslimtarihi} ‘de 3’ÜNCÜ ZIRHLI TUGAY KOMUTANLIĞI'nın Kütüphanesinde bulunan "
         f"{kitapadi}, {kitapno} No’lu kitabı, {birlik} personeli olan {rutbe} {adsoyad}, {telno} "
         "telefon numaralı personel kitabı tarafımızca teslim almıştır. Kitabı kütüphaneye en geç geri "
         f"iade edeceği tarih {iadetarihi} ‘dir."
    )
        paragraph = doc.add_paragraph(text)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.space_after = Pt(12)
        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            run.font.size = Pt(12)

    # Boş satır
        doc.add_paragraph("")
        doc.add_paragraph("")
        doc.add_paragraph("")

    # İmza kısmı (iki sütun gibi görünmesi için tablo kullanabiliriz)
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_ALIGN_PARAGRAPH.LEFT
        table.autofit = False
        table.columns[0].width = Pt(300)
        table.columns[1].width = Pt(300)

        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)

        cell_left_paragraph = cell_left.paragraphs[0]
        cell_left_paragraph.add_run("Kütüphane Personeli / İmza")
        
    # Sağdaki hücreye isim ve rütbe alt alta
        cell_right_paragraph = cell_right.paragraphs[0]
        cell_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cell_right_paragraph.add_run("Teslim Alan / İmza")
       

        # Masaüstüne kaydet
        masaustu = os.path.join(os.path.expanduser("~"), "Desktop")
        dosya_yolu = os.path.join(masaustu, f"{adsoyad}_tutanak.docx")
        doc.save(dosya_yolu)

        QMessageBox.information(None, "Başarılı", f"Tutanak başarıyla oluşturuldu:\n{dosya_yolu}")

    except Exception as e:
        QMessageBox.critical(None, "Hata", str(e))
    finally:
        conn.close()       


ui.pushButton_ekle.clicked.connect(tabloya_satir_ekle)
ui.btnEkle.clicked.connect(EKLE)
ui.btnExel.clicked.connect(EXCEL_AKTAR)
ui.btnTutanak.clicked.connect(TUTANAK)
ui.teslimguncelle.clicked.connect(GUNCELLEemanet)
ui.btnListele.clicked.connect(LISTELE)
ui.btnCikis.clicked.connect(CIKIS)
ui.btnGeri.clicked.connect(geriMenu)
ui.btnSil.clicked.connect(SIL)
ui.silbutonu.clicked.connect(IADE)
ui.btnAd.clicked.connect(KITAPARA)
ui.btnYazar.clicked.connect(YAZARARA)
ui.btnNumara.clicked.connect(NUMARA)
ui.btnDil.clicked.connect(DILARA)
ui.btnTur.clicked.connect(TURARA)
ui.tblwBilgi.itemSelectionChanged.connect(DOLDUR)
ui.emanetler.itemSelectionChanged.connect(DOLDUR2)
ui.btnGuncelle.clicked.connect(GUNCELLE)
ui.menuHakkinda.triggered.connect(HAKKINDA)
ui3.pushButton_2.clicked.connect(kayitiac)    
ui3.sayfalar.clicked.connect(giris)    
ui3.checkBox.clicked.connect(sifre_goster)
ui3.actionKay_t_Ol.triggered.connect(kayitiac)
ui3.action_k.triggered.connect(ANACIKIS)
ui3.actionHakk_nda.triggered.connect(HAKKINDA)
ui4.giris.clicked.connect(girisiac)
ui4.kayitol.clicked.connect(kayitol)
ui5.anasayfaButon.clicked.connect(anasayfa)
ui5.istatistikButon.clicked.connect(istatistikAc)
ui5.hakkindaButon.clicked.connect(hakkindaac)
ui5.cikisButon.clicked.connect(CIKIS2)
ui6.geriButon.clicked.connect(geriIstatistik)


sys.exit(app.exec_())