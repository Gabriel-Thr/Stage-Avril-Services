from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os
from tkinter import PhotoImage
from datetime import datetime
import re
from ftplib import FTP
from configparser import ConfigParser
import pandas as pd
import logging
import logging.config
from Crypto.Cipher import AES

class XML_converter:
    """
    Classe qui convertit un fichier CSV en fichier XML

    Attributs
    ---------
    logger : Logger
        log de la classe
    erreur_date : int
        compte le nombre de date entrée invalide
    erreur_csv : int
        compte le nombre d'erreur dans le CSV
    erreur : int
        somme des erreurs de la classe

    
    Méthode
    -------
    CSVtoXML(self, input_file, output_file, fournisseur, code_fournisseur, societe, code_societe, date_application, date_expiration, date_import)
        prend en entrée le fichier CSV, le nom du fichier XML, la date actuelle ainsi que les données rentrées par l'utilisateur et 
        retourne un fichier XML à partir des données en entrée
    """

    def __init__(self):
        self.logger = logging.getLogger('XML_Converter')
        self.erreur_date = 0
        self.erreur_csv = 0
        self.erreur = self.erreur_date + self.erreur_csv

    def CSVtoXML(self, input_file, output_file, fournisseur, code_fournisseur, societe, code_societe, date_application, date_expiration, date_import):
        if not input_file.endswith('.csv'):
            self.logger.warning("Fichier CSV attendu en entrée")
        if not output_file.endswith('.xml'):
            self.logger.warning("Fichier XML attendu en sortie")

        try:
            df = pd.read_csv(input_file, encoding='unicode_escape', on_bad_lines='warn', sep=r'\s*;\s*', engine='python', header=None, dtype=str)
            self.logger.info("Lecture du fichier CSV réussi")
        except FileNotFoundError:
            self.logger.warning("Fichier CSV introuvable")
            return

        entire_xml = ''
        header_xml = ''
        tarif_xml = ''

        col = df.columns
        #Année entre 2023 et 2099, Mois entre 01 et 12, Jour entre 01 et 31 en fonction du mois, prend en compte toutes les années bissextiles entre 2023 et 2099
        checkDate = r'^(202[3-9]|20[3-9]\d)(0[13578]|1[02])(0[1-9]|[12]\d|3[01])$|^(202[3-9]|20[3-9]\d)(0[469]|11)(0[1-9]|[12]\d|30)$|^(202[3-9]|20[3-9]\d)02(0[1-9]|1\d|2[0-8])$|^(202[48]|20[2468][048]|20[13579][26]|20[02468][048]|20[13579][26])0229$' 
        if not (len(date_application) == 8) or not (re.search(checkDate, date_application)):
            self.logger.warning(f"Format de date d'application invalide : {date_application}")
            self.erreur_date += 1
        elif (len(date_application) == 0):
            self.logger.info("Date d'application non renseignée")
        else:
            self.logger.debug(f"Date d'application valide : {date_application}")
        if (not (len(date_expiration) == 8)) or (not (re.search(checkDate, date_expiration))):
            self.logger.warning(f"Format de date d'expiration invalide : {date_expiration}")
            self.erreur_date += 1
        elif (len(date_expiration) == 0):
            self.logger.info("Date d'expiration non renseignée")
        elif (date_expiration < date_application):
            self.logger.warning(f"Date d'expiration inférieur à la date d'application : {date_expiration} < {date_application}")
            self.erreur_date += 1
        else:
            self.logger.debug(f"Date d'expiration valide : {date_expiration}")
        if len(df.columns) != 5:
            self.logger.error(f'Erreur : {len(df.columns)} colonnes au lieu de 5')
            self.erreur_csv += 1

        header_xml = f"""
    <MessageHeader>
        <Sender>{fournisseur}</Sender>
        <Receiver>{societe}</Receiver>
        <Date>{date_import}</Date>
    </MessageHeader>"""

        for j in range(len(df)):
            code_article = df[col[0]][j]
            conditionnement = df[col[1]][j]
            depart_rendu = df[col[2]][j]
            prix = df[col[3]][j]
            unite = df[col[4]][j]

            prix_str = str(prix) #Pour remplacer la virgule par un point sans erreur

            if j > 0 and (pd.isna(code_article) and pd.isna(conditionnement) and pd.isna(depart_rendu) and pd.isna(prix) and pd.isna(unite)): #si ligne vide
                tarif_xml += ""

            elif j > 0 and (pd.isna(code_article) or pd.isna(conditionnement) or pd.isna(depart_rendu) or pd.isna(prix) or pd.isna(unite)):
                self.logger.error(f'Erreur ligne {j}, donnée non renseignée')
                self.erreur_csv += 1
            else:
                tarif_xml += f"""
        <Tariff>
            <Reference AssignedBy="Sender"></Reference>
            <Product AssignedBy="Receiver">{code_article}</Product>
            <Product AssignedBy="Sender"></Product>
            <Product AssignedBy="Supplier"></Product>
            <Packaging AssignedBy="Receiver">{conditionnement}</Packaging>
            <Packaging AssignedBy="Sender"></Packaging>
            <Packaging AssignedBy="Supplier"></Packaging>
            <Supplier AssignedBy="Receiver">S;{code_fournisseur}</Supplier>
            <Supplier AssignedBy="Sender"></Supplier>
            <Company AssignedBy="Receiver">{code_societe}</Company>
            <Company AssignedBy="Sender"></Company>
            <Place AssignedBy="Receiver"></Place>
            <Place AssignedBy="Sender"></Place>
            <Activity AssignedBy="Receiver"></Activity>
            <Activity AssignedBy="Sender"></Activity>
            <DeliveryType AssignedBy="Receiver">{depart_rendu}</DeliveryType>
            <ApplicationDate>{date_application}</ApplicationDate>
            <ExpiryDate>{date_expiration}</ExpiryDate>
            <Price>{prix_str.replace(',', '.')}</Price>
            <PriceCurrency AssignedBy="Receiver">EUR</PriceCurrency>
            <PriceCurrency AssignedBy="Sender"></PriceCurrency>
            <PriceUom AssignedBy="Receiver">{unite}</PriceUom>
            <PriceUom AssignedBy="Sender"></PriceUom>
            <Remark></Remark>
        </Tariff>
        """
        
        entire_xml = f"""<?xml version="1.0" encoding="iso-8859-1"?>
<Message>{header_xml}    
    <Tariffs>{tarif_xml}       
    </Tariffs>
</Message>"""

        if self.erreur_csv == 0:
            self.logger.info("CSV valide")

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(entire_xml)


class FTPManager:
    """
    Classe qui envoie lit un fichier de conf et qui décrypte son mot de passe et envoie un fichier sur le serveur FTP associé au fichier de conf

    Attribut
    --------
    logger : Logger
        log de la classe
    erreur : int
        compteur du nombre d'erreur se produisant durant l'exécution du code
    
    Méthode
    -------
    load_key(self, file_path):
        Lit la clé permettant de déchiffrer le mot de passe
    decrypt_password(self, config, key):
        Déchiffre le mot de passe à partir de la clé lu au préalable, et du fichier de conf
    sendOnFTP(self, ftp_server, ftp_user, ftp_password, ftp_path, file):
        Prend en entrée les informations du fichier de conf permettant de se connecter au serveur FTP, et établit la connexion au serveur FTP,
        puis ajoute le fichier au serveur et enfin, ferme la connexion
    """

    def __init__(self):
        self.logger = logging.getLogger('FTPManager')
        self.erreur = 0

    def load_key(self, file_path):
        try:
            with open(file_path, "rb") as key_file:
                key = key_file.read()
            self.logger.info("Lecture de la clé réussi")
            return key
        except FileNotFoundError:
            self.logger.warning('Key file not found')
            self.erreur += 1
        except Exception as e:
            self.logger.warning(f'Erreur lors de la lecture du fichier de clé: {e}')
            self.erreur += 1

    def decrypt_password(self, config, key):
        try:
            nonce = bytes.fromhex(config.get("FTP", "Password_nonce"))
            tag = bytes.fromhex(config.get("FTP", "Password_tag"))
            cyphertext = bytes.fromhex(config.get("FTP", "Password_cyphertext"))

            self.logger.debug(f"Nonce: {nonce.hex()}")
            self.logger.debug(f"Tag: {tag.hex()}")
            self.logger.debug(f"Cyphertext: {cyphertext.hex()}")
            self.logger.debug(f"Key: {key.hex()}")

            cypher = AES.new(key, AES.MODE_EAX, nonce=nonce)
            decrypted_password = cypher.decrypt_and_verify(cyphertext, tag)
            self.logger.info("Déchiffrement du mot de passe réussi")
            return decrypted_password.decode()
        except Exception as e:
            self.logger.warning(f"Erreur lors du déchiffrement du mot de passe : {e}")
            self.erreur += 1

    def sendOnFTP(self, ftp_server, ftp_user, ftp_password, ftp_path, file):
        try:
            with FTP(ftp_server) as ftp:
                ftp.login(ftp_user, ftp_password)
                ftp.set_pasv(True)
                ftp.cwd(ftp_path)
                with open(file, 'rb') as f:
                    ftp.storbinary('STOR ' + file, f)
                ftp.close()
                self.logger.info("Envoi du fichier sur le serveur FTP réussi")
        except Exception as e:
            self.logger.warning(f"Erreur pendant l'envoi du fichier : {e}", exc_info=True)
            self.erreur += 1


class ihmOperra:
    """
    Classe représentant l'interface utilisateur pour l'envoi du fichier XML sur le serveur FTP

    Attributs:
    ----------
    root : Tk
        instance de la fenêtre principale de l'application
    converter : XML_converter
        instance de la classe XML_converter pour gérer la conversion des fichiers CSV en XML
    ftp_manager : FTPManager
        instance de la classe FTPManager pour gérer les transferts FTP
    logger : logging.Logger
        log de la classe
    csv_file_path : StringVar
        chemin du fichier CSV sélectionné
    date_import : str
        date et heure actuelles
    key_path : str
        chemin vers le fichier de clé de déchiffrement du mot de passe
    erreur : int
        compteur du nombre d'erreur se produisant durant l'exécution du code
    current_directory : str
        répertoire courant du script permettant de récupérer certains fichier nécessaire (logo et conf_path)
    image_path : str
        chemin vers le fichier de l'image (logo)
    logo : PhotoImage
        image du logo à afficher dans l'interface
    conf_path : str
        chemin vers le fichier de configuration FTP
    date_application_var : StringVar
        variable pour la date d'application entrée par l'utilisateur
    date_expiration_var : StringVar
        variable pour la date d'expiration entrée par l'utilisateur

    Méthodes:
    ---------
    init_ui(self)
        configure et affiche les éléments de l'interface utilisateur
    open_file_dialog(self)
        ouvre une boîte de dialogue pour sélectionner un fichier CSV
    extract_name(self, file_path)
        extrait et retourne le nom de base d'un fichier CSV sans chiffres 
    validate_date(self, *args)
        valide les dates d'application et d'expiration et met à jour leur couleur de texte en fonction de leur validité, rouge si invalide, noir sinon
    clear_interface(self)
        réinitialise tous les champs de l'interface utilisateur
    suppr_file(self, file)
        Supprime le fichier XML local créé
    transformer(self)
        Effectue la transformation du fichier CSV en fichier XML, l'envoie sur le serveur FTP, la supression du fichier XML et l'effaçage de l'interface
    """

    def __init__(self, root):
        self.root = root

        self.converter = XML_converter()
        self.ftp_manager = FTPManager()

        self.logger = logging.getLogger('ihmOperra')
        
        self.csv_file_path = StringVar()
        self.date_import = datetime.now().strftime('%Y,%m,%d %H:%M:%S')
        self.key_path = 'test_ftp_key.key'
        self.erreur = 0

        self.current_directory = os.path.dirname(os.path.abspath(__file__))
        self.image_path = os.path.join(self.current_directory, 'avril.png')
        self.logo = PhotoImage(file=self.image_path)

        self.conf_path = os.path.join(self.current_directory, 'test_ftp_operra_tarif_achat.conf')

        self.date_application_var = StringVar()
        self.date_expiration_var = StringVar()

        self.date_application_var.trace_add('write', self.validate_date)
        self.date_expiration_var.trace_add('write', self.validate_date)

        self.init_ui()

    def init_ui(self):
        self.root.title("Formulaire")
        self.frm = ttk.Frame(self.root, padding=10)
        self.frm.grid()

        ttk.Label(image=self.logo).grid(column=0,row=0)
        ttk.Label(text='Création FICHIER XML:').grid(column=0, row=1)

        ttk.Label(text='Nom fournisseur').grid(column=0, row=2)
        self.fournisseur_entry = ttk.Entry()
        self.fournisseur_entry.grid(column=0, row=3)

        ttk.Label(text='Code fournisseur').grid(column=0, row=4)
        self.code_fournisseur_entry = ttk.Entry()
        self.code_fournisseur_entry.grid(column=0, row=5)

        ttk.Label(text='Société').grid(column=0, row=6)
        self.societe_entry = ttk.Entry()
        self.societe_entry.grid(column=0, row=7)

        ttk.Label(text='Code société').grid(column=0, row=8)
        self.code_societe_entry = ttk.Entry()
        self.code_societe_entry.grid(column=0, row=9)

        ttk.Label(text='Date Application (format YYYYMMJJ)').grid(column=0, row=10)
        self.date_application_entry = ttk.Entry(textvariable= self.date_application_var)
        self.date_application_entry.grid(column=0, row=11)

        ttk.Label(text='Date Expiration (format YYYYMMJJ)').grid(column=0, row=12)
        self.date_expiration_entry = ttk.Entry(textvariable = self.date_expiration_var)
        self.date_expiration_entry.grid(column=0, row=13)

        ttk.Label(text='Fichier CSV').grid(column=0, row=14)

        self.labelCsv = ttk.Label(self.root, text='Aucun fichier sélectionné')
        self.labelCsv.grid(column=0, row=16)
        buttonCSV = ttk.Button(self.root, text='Choisir un fichier', command=self.open_file_dialog)
        buttonCSV.grid(column=0, row=15)

        resButton = ttk.Button(self.root, text='Transformer', command=self.transformer)
        resButton.grid(column=0, row=18)

        self.res = ttk.Label(self.root, text='')
        self.res.grid(column=0, row=19)

        destroyButton = ttk.Button(self.root, text='Fermer', command=self.root.destroy)
        destroyButton.grid(column=0, row=20)

    def open_file_dialog(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.labelCsv.config(text=os.path.basename(filename))
        self.csv_file_path.set(filename)

    def extract_name(self, file_path):
        file_name = file_path.split('/')[-1]
        name = file_name.replace('.csv', '')
        name_for_xml = re.sub(r'\d+', '', name)
        self.logger.info(f"Extraction du nom du fichier : {name_for_xml}")
        return name_for_xml
    
    def validate_date(self, *args):
        def validate(entry, date_var):
            date = date_var.get()
            if re.fullmatch(r'^(202[3-9]|20[3-9]\d)(0[13578]|1[02])(0[1-9]|[12]\d|3[01])$|^(202[3-9]|20[3-9]\d)(0[469]|11)(0[1-9]|[12]\d|30)$|^(202[3-9]|20[3-9]\d)02(0[1-9]|1\d|2[0-8])$|^(202[48]|20[2468][048]|20[13579][26]|20[02468][048]|20[13579][26])0229$', date):
                entry.config(foreground='black')
            else:
                entry.config(foreground='red')

        validate(self.date_application_entry, self.date_application_var)
        validate(self.date_expiration_entry, self.date_expiration_var)

    def clear_interface(self):
        self.fournisseur_entry.delete(0, END)
        self.code_fournisseur_entry.delete(0, END)
        self.societe_entry.delete(0, END)
        self.code_societe_entry.delete(0, END)
        self.date_application_entry.delete(0, END)
        self.date_expiration_entry.delete(0, END)
        self.labelCsv.config(text='Aucun fichier sélectionné')
        self.logger.info("Champs de l'interface effacée")

    def suppr_file(self, file):
        try:
            os.remove(file)
            self.logger.info("Supression du fichier XML local")
        except Exception as e:
            self.logger.error(f'Erreur lors de la suppression du fichier local : {e}', exc_info=True)
            self.erreur += 1

    def transformer(self):
        self.erreur = 0
        csv_file = self.csv_file_path.get()
        xml_file = f'test_OPERRA_{self.extract_name(csv_file)}_{self.date_application_entry.get()}.xml'

        if csv_file and xml_file:
            try:
                self.converter.CSVtoXML(csv_file, xml_file, self.fournisseur_entry.get(), self.code_fournisseur_entry.get(), self.societe_entry.get(), self.code_societe_entry.get(),
                                      self.date_application_entry.get(), self.date_expiration_entry.get(), self.date_import)
                if self.converter.erreur != 0:
                    self.suppr_file(xml_file)
                    self.converter.erreur = 0
                    return self.res.config(text='Erreur pendant la création du fichier XML (voir log)')
                self.logger.info("Convertion du CSV en XML réussi")
                self.res.config(text='Fichier converti')

                config = ConfigParser()
                config.read(self.conf_path)
                ftp_server = config.get('FTP', 'Host')
                ftp_user = config.get('FTP', 'User')
                ftp_password = self.ftp_manager.decrypt_password(config, self.ftp_manager.load_key(self.key_path))
                ftp_path = config.get('FTP', 'FTPPath')

                self.ftp_manager.sendOnFTP(ftp_server, ftp_user, ftp_password, ftp_path, xml_file)
                if self.ftp_manager.erreur == 0:
                    self.res.config(text='Fichier envoyé sur le serveur FTP')
                    self.suppr_file(xml_file)
                    if self.erreur == 0:
                        self.clear_interface()
                        self.logger.info("Processus effectué avec succès\n")
                else:
                    self.res.config(text='Erreur pendant l\'envoi du fichier FTP (voir log)')

            except Exception as e:
                dt = datetime.now().strftime("%d/%m/%y %H:%M:%S")
                self.logger.error(f"Erreur lors de la création du fichier XML: {e}", exc_info=True)
                self.erreur += 1
                return self.res.config(text=f"Erreur (voir log): {e} | {dt}")


if __name__ == "__main__":
    def get_log_level():
        LOG_LEVELS = {
            50: logging.CRITICAL,
            40: logging.ERROR,
            30: logging.WARNING,
            20: logging.INFO,
            10: logging.DEBUG,
        }

        config = ConfigParser()
        config.read(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_ftp_operra_tarif_achat.conf'))

        log_level_str = config.get('Level','log_level',fallback='20')
        log_level = int(log_level_str)

        return LOG_LEVELS.get(log_level)
    
    logging.basicConfig(filename='test_formulaire_tarif_operra.log', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=get_log_level())
    root = Tk()
    ihm = ihmOperra(root)
    root.mainloop()