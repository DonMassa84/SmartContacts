#!/usr/bin/python
# -*- coding: utf-8 -*-
# legt eine Sicherung der Dateien des Rechners an

import sys, string, os, traceback, xml.dom.minidom, glob, shutil, locale, datetime, time, collections #, adb_android
from optparse import OptionParser
#from adb_android import adb_android

def getTextFromXml(element):
  rc = ""
  for node in element.childNodes:
    if node.nodeType == node.TEXT_NODE:
      rc = rc + node.data  #.encode('iso-8859-15')
      
  return rc

#spezielle Exception, die keinen Trace mit ausgeben soll
class GmException( Exception ):
  def __init__( self, Fehler ):
    self.message = Fehler
   
  def __str__( self ):
    return self.message

class SettingsXml:
  """Kapselt den Zugriff auf die Settings.xml"""
  def __init__( self, Dateipath ):
    document = ""
    with open( Dateipath, 'r', encoding="iso-8859-1" ) as SettingsFile: 
      for line in SettingsFile:
        document += line
    self.Xml = xml.dom.minidom.parseString(document)
    options = self.Xml.getElementsByTagName("options")
    if options is None or len( options ) <= 0:
      raise GmException( "Der Tag options ist in der Settings-Datei nicht angegeben." )
    self.Options = options[0]
    self.Rechner = None # wird beim Zugriff auf getPaths gelesen und gefüllt
    self.Typ = None # wird beim Zugriff auf getPaths gelesen und gefüllt
    self.ZielPath = None # wird beim ersten Zugriff geladen

  def getPaths( self ):
    paths = self.Xml.getElementsByTagName("paths")
    if paths is None or len( paths ) <= 0:
      raise GmException( "Der Tag paths ist in der Settings-Datei nicht angegeben." )
    pathsChilds = paths[0].childNodes 
    self.Rechner = paths[0].getAttribute("rechner")
    if self.Rechner is None or len( self.Rechner ) <= 0:
      raise GmException( "Das Attribut rechner wurde im Tag paths der Settings-Datei nicht angegeben." )
    self.Typ = paths[0].getAttribute("typ")
    return pathsChilds

  def getBackupZielPath( self ):
    if self.ZielPath is None:
      self.ZielPath = getTextFromXml(self.Options.getElementsByTagName("ziel")[0])
      if self.ZielPath is None or len( self.ZielPath ) <= 0:
        raise GmException( "Der Tag ziel wurde im Tag options der Settings-Datei nicht angegeben." )

    ZielPath = os.path.join( self.ZielPath, self.Rechner )
    heute = datetime.datetime.now()
    jahr = '{:%Y}'.format(heute)
    ZielPath = os.path.join( ZielPath, jahr )
    VollBackup = os.path.join( ZielPath, "VollBackup" )
    Increment = '{:Inkrement_%Y_%m_%d}'.format(heute)
    ZielPath = os.path.join( ZielPath, Increment )
    return ( VollBackup, ZielPath )
      
class PathSettings( object ):
  def __init__(self, settings, xmlItem):
    self.SettingsXml = settings
    self.Quelle = xmlItem.getAttribute("quelle")
    self.Ziel = xmlItem.getAttribute("ziel")  
    sGroesse = xmlItem.getAttribute( "MaxFilesInZip" )
    try:
      self.MaxGroesse = int( sGroesse )
    except ValueError:
      self.MaxGroesse = 0

  def __relativePath( self, base,path):
    return path.replace(base + os.sep,'',1) 

  def read( self ):    
    rc = 0  
    if self.SettingsXml.Typ == "Smartphone":
      print( adb_android.version() )
      return
    if self.Quelle is None:
      print( "Die Quelle muss angegeben werden." )
      rc = 3
    elif self.Ziel is None:
      print( "Das Zielverzeichnis muss angegeben werden." )
      rc = 3
    elif not os.path.exists( self.Quelle ):
      print( "Die angegebene Quelle existiert nicht:'" + self.Quelle + "'" )
      Dir = os.path.dirname( self.Quelle )
      if len( Dir ) == 0:
        Dir = os.path.curdir
        print( "Dateien im Verzeichnis " + Dir )
        print( os.listdir( Dir ) )
      rc = 3
    if rc != 0:
      return rc

    # zu kopierende Dateien bestimmen
    # alle auslesen
    QuellFiles = self.readQuelle( self.Quelle )

    ( VollBackupPath, ZielPath ) = self.SettingsXml.getBackupZielPath()

    KatalogFile = os.path.join( VollBackupPath, self.Ziel + "_Katalog.xml" )
    Katalog = KatalogXml( KatalogFile )

    BackupFile = os.path.join( VollBackupPath, self.Ziel + ".zip" )
    if os.path.isfile( BackupFile ):
      # Increment bestimmen
      # Katalogfile einlesen - beim Vollbackup steht immer der letzte Stand
      QuellFiles = Katalog.passeGesicherteAn( QuellFiles )

      if len( QuellFiles ) > 0:
        # in incrementelles Backupverzeichnis speichern
        BackupFile = os.path.join( ZielPath , self.Ziel + ".zip" )
        if not os.path.isdir( ZielPath ):
          print( "Das Verzeichnis für die Teilsicherung wird angelegt: " + ZielPath );
          os.makedirs(ZielPath)
          if not os.path.isdir( ZielPath ):
            print( "Das angegebene Zielverzeichnis ist kein Verzeichnis: " + ZielPath )
            rc = 3
    else:
      Katalog.setQuellen( QuellFiles )
      # Vollbackup anlegen
      ZielPath = VollBackupPath
      if not os.path.isdir( VollBackupPath ):
        print( "Das Verzeichnis für die Vollsicherung wird angelegt: " + VollBackupPath );
        os.makedirs(VollBackupPath)
        if not os.path.isdir( VollBackupPath ):
          print( "Das angegebene Zielverzeichnis ist kein Verzeichnis: " + VollBackupPath )
          rc = 3
    if rc != 0:
      return rc

    if len( QuellFiles ) > 0:
      # Datei schreiben, die die Dateien enthält, die kopiert werden sollen
      Index = 0
      ZipNr = 0
      BackupListFile = os.path.join( ZielPath, self.Ziel + "_Zip.txt" )
      ZipListe = open( BackupListFile, 'w', encoding="UTF-8", errors="strict" )
      for QuellFile in sorted( QuellFiles ):
        Index = Index + 1
        ZipListe.write( QuellFile + "\n" )
        if self.MaxGroesse > 0 and Index >= self.MaxGroesse:
          ZipListe.close()
          ZipNr = self.compress( ZipNr, BackupListFile, BackupFile )
          Index = 0
          ZipListe = open( BackupListFile, 'w', encoding="UTF-8", errors="strict" )


      ZipListe.close()
      # Verzeichnis packen und ins Zielverzeichnis schreiben
      if Index > 0:
        self.compress( ZipNr, BackupListFile, BackupFile )

      # Katalogdatei schreiben
      Katalog.writeKatalog()

    return rc

  def readQuelle( self, src, funcSollFileKopiertWerden = None, WithSubDirectories = True ):
    print( "Lese " + src )
    QuellFiles = {}
    for root, folders, files in os.walk(src,topdown=True):
      if not WithSubDirectories:
        del folders[:]
      for filename in files:
        srcpath = os.path.join(root,filename)
        try:   
          #print( srcpath )
          #if funcSollFileKopiertWerden == None or funcSollFileKopiertWerden(dstpath):

          # letztes Änderungsdatum auslesen, in lokale Zeit umrechnen und formatieren
          aenderung = time.strftime( '%d.%m.%Y %H:%M:%S', time.localtime( os.path.getmtime( srcpath ) ) )
          # Größe auslesen
          groesse = os.path.getsize( srcpath )
          QuellFiles[ self.__relativePath( src, srcpath ) ] = ( aenderung, groesse )
        except BaseException as Ex:
          print( "Datei nicht lesbar: " + srcpath )
          print( Ex )
    return QuellFiles

  def compress( self, ZipNr, Verzeichnisliste, ErgebnisFileParam ):
    ErgebnisFile = ErgebnisFileParam
    while ( os.path.isfile(ErgebnisFile) ):
      ZipNr = ZipNr + 1
      ErgebnisFile = ErgebnisFileParam.replace( ".zip", "_" + str( ZipNr) + ".zip" )
      
    print( "Erstelle " + ErgebnisFile )
    os.chdir(self.Quelle)
    # a  Add
    # -ssw nimmt auch Dateien, die geöffnet sind
    # -tzip erstellt Zip-File
    rc = os.system( r'"C:\Program Files\7-Zip\7z.exe" a -ssw -tzip ' + ErgebnisFile + " @" + Verzeichnisliste )
    if ( rc > 1 ):
      raise GmException( "Fehler bei der Sicherung/Komprimierung von " + ErgebnisFile )
    return ZipNr


# ****************************************************************************
# ********* class KatalogXml ************************************************
# ****************************************************************************
  
class KatalogXml:
  """Kapselt den Zugriff auf die Katalogdatei eines Verzeichnisses"""
  def __init__( self, Dateipath ):
    self.Dateipath = Dateipath
    if os.path.isfile( Dateipath ):
      document = ""
      with open( Dateipath, 'r', encoding="iso-8859-1" ) as SettingsFile: 
        for line in SettingsFile:
          document += line
      self.Xml = xml.dom.minidom.parseString(document)
    else:
      self.Xml = None
    self.GesicherteFiles = None
      
  def readGesicherteFiles( self ):
    if self.GesicherteFiles is None:
      self.GesicherteFiles = {}
      if self.Xml is not None:
        self.readKatalog()
    return self.GesicherteFiles

  def setQuellen( self, QuellFiles ):
    self.GesicherteFiles = QuellFiles

  def passeGesicherteAn( self, QuellFiles ):
    ErgebnisFiles = {}
    self.readGesicherteFiles()

    for Quelle in QuellFiles:     
      ( aenderungQuelle, groesseQuelle ) = QuellFiles[ Quelle ]
      if Quelle in self.GesicherteFiles:
        # Datei existiert bereits in Sicherung
        ( aenderungSicherung, groesseSicherung ) = self.GesicherteFiles[ Quelle ]
        if aenderungQuelle != aenderungSicherung or groesseQuelle != groesseSicherung:
          # Datei seit letzter Sicherung geändert
          ErgebnisFiles[ Quelle ] = ( aenderungQuelle, groesseQuelle )
          self.GesicherteFiles[ Quelle ] = ( aenderungQuelle, groesseQuelle )
      else:
        # Datei existiert noch nicht in Sicherung
          ErgebnisFiles[ Quelle ] = ( aenderungQuelle, groesseQuelle )
          self.GesicherteFiles[ Quelle ] = ( aenderungQuelle, groesseQuelle )
    return ErgebnisFiles

  def readKatalog( self ):
    files = self.Xml.getElementsByTagName("FileList")
    if files is None or len( files ) <= 0:
      print( self.Dateipath + " enthält keine Dateien." )
    filesChilds = files[0].childNodes 
    for item in filesChilds:
      if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE and item.tagName == "File":
        aenderung = item.getAttribute( "LastTime" )
        sGroesse = item.getAttribute( "Size" )
        try:
          groesse = int( sGroesse )
        except ValueError:
          raise GmException( "Größe ist keine Zahl: '" + sGroesse + "' bei " + self.Dateipath )
        filename = getTextFromXml( item )
        self.GesicherteFiles[ filename ] = ( aenderung, groesse )

  def writeKatalog( self ):
    impl = xml.dom.minidom.getDOMImplementation()
    self.Xml = impl.createDocument(None, "FileList", None)
    for Datei in sorted( self.GesicherteFiles ):
      DateiElement = self.Xml.createElement("File")
      self.Xml.documentElement.appendChild(DateiElement)
      TextElement = self.Xml.createTextNode(Datei)
      DateiElement.appendChild(TextElement)
      ( aenderung, groesse ) = self.GesicherteFiles[ Datei ]
      DateiElement.setAttribute( "LastTime", aenderung )
      DateiElement.setAttribute( "Size", str( groesse ) )


    s = self.Xml.toprettyxml(indent="  ", newl="\n", encoding = "iso-8859-1")
    with open( self.Dateipath, 'w', encoding="iso-8859-1" ) as KatalogFile:
      KatalogFile.write(s.decode('iso-8859-15'))
    
if __name__ == '__main__':
  print( 'Sichert die Dateien des Rechners' )

  parser = OptionParser()
  parser.add_option( '-s', '--settings', help='XML-Datei mit den Einstellungen')

  rc = 0
  try:
    (options, args ) = parser.parse_args()
  
    if options.settings is None:
      print( "Die Datei mit den Einstellungen muss angegeben werden." )
      rc = 1
    elif not os.path.exists( options.settings ):
      print( "Die angegebene Einstellungsdatei existiert nicht:'" + options.settings + "'" )
      Dir = os.path.dirname( options.settings )
      if len( Dir ) == 0:
        Dir = os.path.curdir
        print( "Dateien im Verzeichnis " + Dir )
        print( os.listdir( Dir ) )
      rc = 2
    elif not os.path.isfile( options.settings ):
      print( "Die angegebene Einstellungsdatei ist keine Datei: " + options.settings )
      rc = 2
    elif not options.settings.endswith( ".xml" ):
      print( "Die angegebene Einstellungsdatei hat nicht die Endung xml: " + options.settings )
      rc = 2

    if rc == 0:
      CurDir = os.getcwd()
      try:
        # Monatsnamen in richtiger Sprache
        locale.setlocale(locale.LC_ALL, '')
      
        #Converter = Converter( options.quelle, options.ziel )
        #rc = Converter.start( options )

        settings = SettingsXml( options.settings )

        for item in settings.getPaths():
          if item.nodeType == xml.dom.minidom.Node.ELEMENT_NODE and item.tagName == "item":
            pathSettings = PathSettings( settings, item )
            rc = pathSettings.read()
            if rc > 0:
              break


      except GmException as Ex:
        print( "" )
        print( "Fehler:" )
        print( Ex )
        rc = 2
      except BaseException as Ex:
        print( "Es ist eine Exception aufgetreten." )
        traceback.print_exc()
        print( "" )
        print( "Fehler:" )
        print( Ex )
        rc = 2
    os.chdir(CurDir)
  except BaseException as Ex:
    rc = 1
    print( Ex )
 
  if rc == 1:  # Parameterfehler    
    print( parser.print_help() )

  if rc != 0:
    print( "\nEs sind Fehler aufgetreten.\n" )
  else:
    print( "\nFehlerfrei beendet.\n" )

  sys.exit( rc )
