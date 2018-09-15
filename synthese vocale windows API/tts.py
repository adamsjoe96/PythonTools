# -*- coding: utf-8 -*-
import win32com.client, struct

# ---------------------------------------------------------------------
#       Les constantes dÃ©finies dans la doc de SAPI 5.1
#       (une classe correspond Ã  un type enum)
# ---------------------------------------------------------------------
class SpeechVoicePriority: # SpVoice.Priority
    SVPNormal = 0
    SVPAlert = 1
    SVPOver = 2

class SpeechVoiceSpeakFlags: # SpVoice.Speak()
    # SpVoice flags
    SVSFDefault = 0
    SVSFlagsAsync = 1
    SVSFPurgeBeforeSpeak = 2
    SVSFIsFilename = 4
    SVSFIsXML = 8
    SVSFIsNotXML = 16
    SVSFPersistXML = 32
    # Normalizer flags
    SVSFNLPSpeakPunc = 64
    # Masks
    SVSFNLPMask = 64
    SVSFVoiceMask = 127
    SVSFUnusedFlags = -128

class SpeechStreamFileMode: # SpFileStream.Open()
    SSFMOpenForRead = 0
    SSFMCreateForWrite = 3

class SpeechRunState: # SpVoice.Status
    SRSEDone = 1
    SRSEIsSpeaking = 2

class sapi_error_codes: # codes d'erreur renvoyÃ©s par sapi
    SPERR_DEVICE_BUSY = -2147201018 # The wave device is busy


# ---------------------------------------------------------------------
#       L'objet SpVoice
# ---------------------------------------------------------------------
class Voix:
    """Cette classe permet de dÃ©finir une voix.

        ParamÃ¨tres :
            <la_voix> : l'objet "voix" retournÃ© par GetVoices(). La valeur par dÃ©faut None
            permet de crÃ©er l'objet sans sÃ©lectionner de voix et d'utiliser la
            voix par dÃ©faut.
            <rate> : dÃ©bit du discours [-10 ; 10] (0 par dÃ©faut)
            <volume> : volume [0 ; 100] (50 par dÃ©faut)
            <media> : mÃ©dia de sortie de la voix (objet). La valeur par dÃ©faut None permet
            d'utiliser le mÃ©dia par dÃ©faut (en gÃ©nÃ©ral la carte son).
            <priorite> : prioritÃ© de la voix (0 : normale par dÃ©faut,
            1 : alerte, 2 : prioritaire)
            <xml> : True (par dÃ©faut) si le texte contient des balises xml devant Ãªtre
            interpÃ©tÃ©es (ex : <pron ...>) ; False sinon
            <async> : False (valeur par dÃ©faut) pour un discours asynchrone ; True
            sinon.
            <evenement> : False par dÃ©faut, si la voix n'Ã©coute pas les Ã©vÃ¨nements ; True
            sinon. La gestion des Ã©vÃ¨nements est valable seulement en mode
            asynchrone.
            <gestion_evenement> : nom d'une classe permettant de gÃ©rer les Ã©vÃ¨nements. La classe
            doit exister dans le script principal. La valeur par dÃ©faut est None. Si <evenement> = True,
            alors <gestion_evenement> doit Ãªtre diffÃ©rent de None. Sinon, <gestion_evenement> est None.
            <timeout> : si une voix synchrone essaie de parler en mÃªme temps qu'une autre
            voix, alors le programme est interrompu. Le timeout (en millisecondes) permet
            Ã  la voix d'attendre que le canal de sortie soit libÃ©rÃ©. Si ce n'est pas le cas
            Ã  l'expiration du dÃ©lai, alors une erreur est retournÃ©e. La valeur par
            dÃ©faut est nulle.
        Remarques : suite aux tests
            1) La base de registre permet de retrouver les attributs d'un
            objet (ex : l'attribut "name" est la description de la voix). Cependant,
            ma base de registre ne correspond pas toujours Ã  la documentation. Pour
            cette raison, les attributs ne sont pas utilisÃ©, mais l'objet est
            directement passÃ© en paramÃ¨tre (ex : <media>).
            2) La gestion des Ã©vÃ¨nements de la voix occupent de la mÃ©moire vive
            supplÃ©mentaire, mais cela n'est pas significatif.
            3) La fonctionnalitÃ© de timeout s'active seulement pour une voix
            synchrone qui essaie de "parler" (mÃ©thode Speak()) alors que le
            canal de sortie est occupÃ© par une autre voix. Lorsque cet autre voix
            a terminÃ©, alors la 1Ã¨re voix "parle" mÃªme si le timeout n'est pas
            atteint.
            4) Un timeout nÃ©gatif semble correspondre Ã  un timeout infini. Dans ce
            cas, une voix synchrone attend que toutes les voix aient parlÃ©. A
            manipuler avec prÃ©cautions car la doc est muette sur le sujet.
            5) La fonctionnalitÃ© timeout est sans effet lorsque le canal de sortie
            est occupÃ© par une autre application (ex : musique internet ou
            fichier .wav). Dans ce cas, les sons sont jouÃ©s simultanÃ©ment."""

    def __init__(self, la_voix = None, rate = 0, volume = 50, media = None, priorite = 0,xml = True,  sync=False, evenement = False, gestion_evenement = None, timeout = 0):
        if evenement == True: # cf doc paramÃ¨tre <evenement> et <gestion_evenement>
            if sync == False:
                print("\nErreur : cette voix n'est pas crÃ©Ã©e car la gestion des Ã©vÃ¨nements en mode synchrone n'est pas possible.\n")
                return None
            else:
                if gestion_evenement == None:
                    print("\nErreur : cette voix n'est pas crÃ©Ã©e car le gestionnaire d'Ã©vÃ¨nement (classe) n'existe pas.\n")
                    return None

        # crÃ©ation de l'objet SpVoice
        if evenement == False:
            self.obj = win32com.client.Dispatch("SAPI.SpVoice")
        else:
            self.obj = win32com.client.DispatchWithEvents("SAPI.SpVoice", gestion_evenement)

        # dÃ©finit les propriÃ©tÃ©s de l'objet SpVoice provenant de l'API
        if la_voix != None:
            self.obj.Voice = la_voix
        self.obj.Rate = rate
        self.obj.Volume = volume
        self.obj.Priority = priorite
        if media != None:
            self.obj.AudioOutput = media

        if sync == False: # initialise le timeout
            self.obj.SynchronousSpeakTimeout = timeout

        # dÃ©finit les flags de la mÃ©thode Speak() (mÃ©thode parler())
        self.speak_flag = SpeechVoiceSpeakFlags.SVSFDefault
        if xml == False:
            self.speak_flag = SpeechVoiceSpeakFlags.SVSFIsNotXML # les balises xml sont prononcÃ©es
        else:
            self.speak_flag = 8 # les balises xml sont interprÃ©tÃ©es
        if sync == True:
            self.speak_flag = self.speak_flag + SpeechVoiceSpeakFlags.SVSFlagsAsync

        # dÃ©finir l'id de la langue de la voix (utilisÃ© pour le lexique)
        self.langue_id = self.__definir_langue_id(la_voix)

    def __definir_langue_id(self, vx = None):
        """Retourne le code langue d'une voix.

            ParamÃ¨tre :
                <vx> : l'objet voix. Si None, alors il est nÃ©cessaire de
                retrouver la voix par dÃ©faut pour obtenir la langue.
            Valeur de retour :
                Le code (dÃ©cimal) de la langue. Si la langue ne peut Ãªtre
                dÃ©terminÃ©e, alors la fonction renvoie None.

            Remarque :
                Les voix disponibles sont dans la catÃ©gorie 'Voices' accessible
                avec la clÃ© de registre
                'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices' (cf doc).
                Les langues sont accessibles avec l'attribut 'Language' qui donne
                une liste de code hÃ©xadÃ©cimal sÃ©parÃ©s par un ';' (cf doc)."""

        voix = vx
        if voix == None: # dÃ©finir la voix par dÃ©faut
            category = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
            category.SetId("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices") # catÃ©gorie Voices
            les_voix = category.EnumerateTokens()
            id_voix_par_defaut = category.Default # id de la voix par dÃ©faut (clÃ© du registre)
            for i in range(0, len(les_voix)):
                if les_voix.Item(i).Id == id_voix_par_defaut:
                    voix = les_voix.Item(i)
                    break

        # dÃ©finir le code langue
        try:
            list_lang = voix.GetAttribute("Language").split(";") # liste des code hÃ©xa du langage
        except:
            print("\nAttention : la langue ne peut Ãªtre dÃ©terminÃ©e pour cette voix\n")
            return None
        else:
            id_lang = int(list_lang[0], 16) # conversion du code hÃ©xa en base 10
            return id_lang

    def __definir_discours(self, discours):
        """Retourne une chaÃ®ne qui sera "parlÃ©e" par le moteur TTS.
            Si le lexique utilisateur contient des mots, alors ils sont
            encadrÃ©s par des sÃ©parateurs afin d'Ãªtre correctement
            prononcÃ©s.

            ParamÃ¨tre :
                <discours> : s'il s'agit d'un nom de fichier, alors le contenu
                total devient la chaÃ®ne Ã  traiter (encodage attendu : utf-8).
                La documentation de la fonction analyser_le_texte() explique
                cette transformation.
                Sinon, le paramÃ¨tre doit Ãªtre une chaÃ®ne.
            Renvoie :
                La chaÃ®ne qui sera "parlÃ©e" par le moteur TTS.
            Remarque :
                On suppose que la mÃ©moire vive est suffisante pour stocker le
                contenu total du fichier."""

        try:
            f = open(discours,mode = "r",encoding = "utf_8_sig",errors = "strict")
        except IOError: # <discours> est une chaÃ®ne
            texte = discours
        else: # <discours> est un fichier
            texte = f.read()
            f.close()

        return analyser_le_texte(texte, self.langue_id)
        # return texte

    def clavier(self):
        """Demande la saisie d'un texte au clavier, et
        demande Ã  la voix de le prononcer."""
        try:
            str = input("Le texte Ã  Ã©couter : ")
        except EOFError:
            pass
        else:
            self.obj.Speak(str)

    def parler(self, discours):
        """La voix prononce un discours.

            ParamÃ¨tres :
                <discours> : une chaÃ®ne ou un chemin de fichier, dont l'encodage
                attendu est utf_8.
            Remarques : suite aux tests
                1) La mÃ©thode Speak() peut lire un fichier directement, mais certains
                caractÃ¨res (ex : Ãª) ne sont alors pas reconnus. Ces caractÃ¨res
                sont correctement traitÃ©s lorsqu'ils font partie d'une chaÃ®ne. Pour
                cette raison, le contenu du fichier est lu puis passÃ© Ã  la
                mÃ©thode Speak() en tant que chaÃ®ne.
                2) La ponctuation (SVSFNLPSpeakPunc) est Ã©noncÃ©e par la voix
                anglaise (ex : '.' devient period), mais pas par les autres
                voix.
                3) La mÃ©thode Speak() retourne toujours 1 mÃªme en mode asynchrone."""

        str = self.__definir_discours(discours)
        try:
            self.obj.Speak(str, self.speak_flag)
        except Exception as erreur: # le timeout est Ã©coulÃ© (voix synchrone)
            print("\nERREUR dans la mÃ©thode .parler() :")
            # recherche le code erreur SPERR_DEVICE_BUSY dans la liste des arguments convertis en chaÃ®ne
            code_erreur = repr(sapi_error_codes.SPERR_DEVICE_BUSY)
            msg = repr(erreur.args)
            if msg.find(code_erreur) != -1:
                print("    Le canal de sortie est occupÃ© par une autre voix")
            else: # erreur inconnue
                print("    Erreur inconnue")
            print("Les 40 premiers caractÃ¨res du texte sont :")
            print("    " + discours[:40])


    def stop(self):
        """ArrÃªter le discours."""
        self.obj.Pause()

    def enregistrer(self, discours, fichier):
        """Enregistrer un discours dans un fichier au format wav.

            ParamÃ¨tres :
                <discours> : le texte ou nom de fichier contenant le discours.
                <fichier> : nom du fichier oÃ¹ est enregistrÃ© la voix.

            Remarques :
                1)La mÃ©thode Speak() doit Ãªtre synchrone, sinon la
                voix n'est pas enregistrÃ©e dans le fichier.
                2) Si le fichier existe dÃ©jÃ  alors il est Ã©crasÃ©.

            Remarque : ajouter le discours au fichier existant ?
                1) Ouverture d'un fichier .wav en mode ajout : KO.
                2) CrÃ©er un fichier .wav temporaire avec la suite du
                discours. Puis ajouter le contenu au fichier initial .wav.
                Ce fichier initial peut Ãªtre lu dans Audacity Ã  condition
                d'Ãªtre importÃ© au format raw (Projet>Importer les donnÃ©es
                brutes -> dans la popup : frÃ©q d'echantillonnage = 22050. Il
                existe nÃ©anmoins une coupure entre le discours et la suite, mais
                Audacity permet de faire un montage.
                3) Cf fonction concatener_wave()."""

        print("DÃ©but de l'enregistrement")
        texte = self.__definir_discours(discours)

        file_stream = win32com.client.Dispatch("SAPI.SpFileStream")
        file_stream.Open(fichier, SpeechStreamFileMode.SSFMCreateForWrite)
        self.obj.AudioOutputStream  = file_stream
        self.obj.Speak(texte)
        file_stream.Close()
        self.obj.AudioOutputStream  = None # la prochaine "voix" sera dirigÃ©e vers le mÃ©dia par dÃ©faut
        print("Fin de l'enregistrement")

    def en_cours(self):
        """Retourne True si la parole est en cours. Sinon, False.

            Remarque :
                Si la voix est synchrone, alors RunningState retourne
                toujours SpeechRunState.SRSEDone (cf doc SAPI 5.1)"""

        if self.obj.Status.RunningState == SpeechRunState.SRSEDone:
            return False
        else:
            return True



# ---------------------------------------------------------------------
#       Les fonctions diverses
# ---------------------------------------------------------------------
def lister_les_voix(selection):
    """Afficher la listes voix installÃ©es, et permet (Ã©ventuellement) d'en
    sÃ©lectionner une. Si la sÃ©lection est valide, alors l'objet "voix" est
    renvoyÃ©. Dans tous les autres cas, la fonction retourne None.

        ParamÃ¨tre :
            <selection> : si True, alors il est possible de sÃ©lectionner une voix."""

    voix = win32com.client.Dispatch("SAPI.SpVoice")
    les_voix = voix.GetVoices()
    nb_voix = les_voix.Count
    voix = None
    print()
    for i in range(0, nb_voix):
        print("(" + str(i) + ") : " + les_voix.Item(i).GetDescription())
    if selection == True:
        print()
        try:
            num = int(input("Quelle est la voix sÃ©lectionnÃ©e ? "))
        except EOFError:
            print("Cette voix n'existe pas")
        except ValueError:
            print("Cette voix n'existe pas")
        else:
            if num < 0 or num >= nb_voix: # erreur Ã  la sÃ©lection
                print("Cette voix n'existe pas")
            else:
                voix = les_voix.Item(num)

    return voix

def lister_les_medias(selection):
    """Lister les mÃ©dias de sortie du son et permet (Ã©ventuellement) d'en
        sÃ©lectionner un. Si la sÃ©lection est valide, alors l'objet "media" est
        renvoyÃ©. Dans tous les autres cas, la fonction retourne None.

        ParamÃ¨tre :
            <selection> : si True, alors la fonction permet de
            choisir le mÃ©dia, et le retourne sous forme d'objet.
            Si le mÃ©dia n'est pas valide ou si <selection> Ã©gale
            False, alors retourne None."""

    media = None
    voix = win32com.client.Dispatch("SAPI.SpVoice")
    les_medias = voix.GetAudioOutputs()
    nb_medias = les_medias.Count
    print()
    for i in range(0, nb_medias):
        print("(" + str(i) + ") : " + les_medias.Item(i).GetDescription())
    if selection == True:
        print()
        try:
            num = int(input("Quelle est le mÃ©dia de sortie sÃ©lectionnÃ© ? "))
        except EOFError:
            print("Ce mÃ©dia n'existe pas")
        except ValueError:
            print("Ce mÃ©dia n'existe pas")
        else:
            if num < 0 or num >= nb_medias: # erreur Ã  la sÃ©lection
                print("Cette mÃ©dia n'existe pas")
            else:
                media = les_medias.Item(num)

    return media

def analyser_le_texte(texte, langue_id):
    """Retourne le texte oÃ¹, les mots du lexique utilisateur ont
        Ã©tÃ© remplacÃ©s par leur Ã©quivalent en majuscules, et encadrÃ©s
        par des espaces pour les sÃ©parer. Cela permet au moteur TTS
        de les identifier, et de leur appliquer la prononciation
        dÃ©finie dans le lexique.

        ParamÃ¨tres :
            <texte> : le texte
            <langue_id> : le code du langage (dÃ©cimal)
        Renvoie :
            le texte Ã  l'identique si le lexique est vide ;
            sinon, le texte modifiÃ©.
        Remarque : suite aux tests
            1) Les mots du lexique sont automatiquement encadrÃ©s
            par des espaces (separateur). Il est possible d'avoir plusieurs
            espaces successifs, mais cela ne gÃªne pas la prononciation.
            2) La transformation en minuscules de la totalitÃ© du texte entraÃ®ne
            la perte de ponctuation, mÃªme en traitant les caractÃ¨res convertibles
            un par un (mystÃ¨re !). Cela ne se produit pas avec les majuscules.
            Donc, les mots sont enregistrÃ©s en majuscules dans le lexique."""

    lexique = lexique = win32com.client.Dispatch("SAPI.SpLexicon")
    tuple = lexique.GetWords()

    les_mots = tuple[0] # la liste des mots (dÃ©jÃ  en majuscules) du lexique
    tuple = None
    nb_mots = les_mots.Count
    if nb_mots == 0:
        return texte

    # crÃ©er la liste des mots qui correspondent Ã  la langue
    list_mots = []
    for i in range(0, nb_mots):
        if les_mots.Item(i).LangId == langue_id:
            list_mots.append(les_mots.Item(i).Word)

    les_mots = None
    nb_mots = len(list_mots)
    if nb_mots == 0:
        return texte

    str = texte.upper() # cf remarque 2)
    str = str.lstrip(" ")
    str = str.rstrip(" ")

    for i in range(0, nb_mots):
        # mot = les_mots.Item(i).Word
        mot = list_mots[i]
        debut = 0
        separateur = " "
        lg_mot = len(mot)
        # pdb.set_trace()
        while bool(1) == True:
            position = str.find(mot, debut)
            if position != -1:
                str = str[0 : position] + separateur + mot + separateur + str[position + lg_mot : ]
                debut = position + lg_mot + 2 # +2 pour tenir compte des sÃ©parateurs
            else:
                break

    return str


# ---------------------------------------------------------------------
#       Les fonctions pour manipuler les fichiers wave (.wav)
# ---------------------------------------------------------------------
def concatener_wave(self, cible, source1, source2):
    """ConcatÃ©ner 2 fichiers wave dans un fichier cible.

        On suppose que les 2 fichiers sont gÃ©nÃ©rÃ©s par la mÃªme application
        et disposent des mÃªmes paramÃ¨tres sonores.
        On suppose aussi que la mÃ©moire vive est suffisante pour conserver
        le + gros des 2 fichiers avant d'Ã©crire le fichier cible.
        On suppose qu'un fichier wave est constituÃ© par
        la suite d'octets :
        Header :
            1 Ã  4 : "RIFF"
            5 Ã  8 : nb d'octets suivants le 9Ã¨me (a)
            9 Ã  12 : "WAVE"
        Chunk : contient les paramÃ¨tres sonores (son, ...)
            Header :
                13 Ã  16 : identifier = "fmt "
                17 Ã  20 : nb d'octets Ã  partir du 20Ã¨me octet constituant le bloc (i)
            Bloc :
                (i) octets
        Chunk : contient les sons
            Header :
                4 octets : identifier = "data"
                4 octets : nb d'octets constituant le bloc (ii)
            Bloc :
                (ii) octets ( = bloc1)

        ConsidÃ©rons un second fichier avec (iii) octets dans le bloc appartenant au
        "chunk" contenant les sons  (= bloc2), dont l'identifier est "data".

        ConcatÃ©ner les 2 fichiers revient donc Ã  crÃ©er un fichier tel que :
        Header :
            1 Ã  4 : "RIFF"
            5 Ã  8 : (a) + (ii)
            9 Ã  12 : "WAVE"
        Chunk :
            Header :
                13 Ã  16 : identifier = "fmt "
                17 Ã  20 : nb d'octets Ã  partir du 20Ã¨me constituant le bloc (i)
            Bloc :
                (i) octets
        Chunk :
            Header :
                4 octets : identifier = "data"
                4 octets : (ii) + (iii)
            Bloc :
                (ii) + (iii) octets (bloc1 + bloc2)

        D'aprÃ¨s les tests, un fichier .wav est constituÃ© par :
            Header
            1 Chunk (identifier = "fmt " <=> paramÃ¨tres sonores)
            1 Chunk (identifier = "data" <=> les sons)

        Les autres mÃ©thodes de concatÃ©nation ne sont pas
        statisfaisantes (cf mÃ©thode Enregistrer()).
    """

    f1 = open(source1, mode = "rb")
    # RIFF header
    riff_header = f1.read(4)
    taille = f1.read(4)
    taille_source1 = struct.unpack_from("L", taille)[0]
    wave_header = f1.read(4)
    # Chunk : identifier = "fmt "
    chunk_source1_fmt_header = f1.read(4)
    chunk_source1_fmt_taille = f1.read(4)
    taille = struct.unpack_from("L", chunk_source1_fmt_taille)[0]
    chunk_source1_fmt_bytes = f1.read(taille)
    # Chunk : identifier = "data"
    chunk_source1_data_header = f1.read(4)
    chunk_source1_data_taille = struct.unpack_from("L", f1.read(4))[0]

    f2 = open(source2, mode = "rb")
    # sauter le RIFF header et le bloc "WAVE"
    f2.seek(12, 0)
    # sauter chunk tel que identifier = "fmt "
    f2.seek(4, 1)
    taille = struct.unpack_from("L", f2.read(4))[0]
    f2.seek(taille, 1)
    # chunk tel que identifier = "data" : lire le bloc
    f2.seek(4, 1)
    chunk_source2_data_taille = struct.unpack_from("L", f2.read(4))[0]

    f = open(cible, mode = "wb")
    f.write(riff_header)
    f.write(struct.pack("L", taille_source1 + chunk_source2_data_taille))
    f.write(wave_header)

    f.write(chunk_source1_fmt_header)
    f.write(chunk_source1_fmt_taille)
    f.write(chunk_source1_fmt_bytes)

    f.write(chunk_source1_data_header)
    f.write(struct.pack("L", chunk_source1_data_taille + chunk_source2_data_taille))
    chunk_data_bytes = f1.read(chunk_source1_data_taille)
    f.write(chunk_data_bytes)
    chunk_data_bytes = f2.read(chunk_source2_data_taille)
    f.write(chunk_data_bytes)
    f1.close()
    f2.close()
    f.close()

def tester_format_wave(nom_fichier):
    """Retourne True si le fichier est au format wave ; False sinon."""

    print()
    print("TESTER LE FORMAT WAVE : " + nom_fichier)
    print()
    f = open(nom_fichier, mode = "rb")

    # calcule la taille rÃ©Ã¨lle
    f.seek(0,2)
    taille_reelle = f.tell()
    f.seek(0, 0)

    # RIFF header
    # 8 octets : -> 1 Ã  4 = "RIFF"
    #            -> 5 Ã  8 = taille du fichier hors header
    octets = f.read(4)
    if str(octets, "ascii", "strict") != "RIFF":
        return False
    octets = f.read(4)
    taille_logique = struct.unpack_from("L", octets)[0] + 8

    if taille_logique != taille_reelle:
        print("attention : taille logique (" + repr(taille_logique) + ") <> taille rÃ©elle (" + repr(taille_reelle) + ")")

    # octets de 9 Ã  12 = "WAVE"
    octets = f.read(4)
    if str(octets, "ascii", "strict") != "WAVE":
        return False

    # La suite du fichier est une suite de "chunk"
    # chunk :
    #   header (8 octets) :
    #       1 Ã  4 : "fmt" ou "data"
    #       5 Ã  8 : nb d'octets aprÃ¨s le header (i) <=> taille du "chunk"
    #   main ((i) octets)
    position = f.tell()
    while position < taille_reelle:
        print("Chunk")
        print("-----")
        print("    header :")
        octets = f.read(4)
        print("        identifier = " + str(octets, "ascii", "strict"))
        octets = f.read(4)
        taille = struct.unpack_from("L", octets)[0]
        print("        taille = " + repr(taille))
        f.seek(taille, 1) # dÃ©placement de taille octets depuis la position courante
        position = f.tell()
        print("position = " + repr(position))

    f.close()
    return True



# ********************************************************
#       corps principal
# ********************************************************

if __name__ == "__main__": # le module est exÃ©cutÃ© indÃ©pendament
    pass