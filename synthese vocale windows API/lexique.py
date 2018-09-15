# -*- coding: utf-8 -*-

import win32com.client 

# ---------------------------------------------------------------------
#       Les constantes dÃ©finies dans la doc de SAPI 5.1
#       (une classe correspond Ã  un type enum)
# ---------------------------------------------------------------------            
class SpeechPartOfSpeech: # SpLexicon.AddPronunciation()
    SPSNotOverriden = -1
    SPSUnknown = 0
    SPSNoun = 4096
    SPSVerb = 8192
    SPSModifier = 12288
    SPSFunction = 16384
    SPSInterjection = 20480
    
# ---------------------------------------------------------------------
#       Les fonctions pour manipuler le lexique (objet SpLexicon)
# ---------------------------------------------------------------------
def dict_code_langue(type):
    """Retourne le dictionnaire (clÃ© = "code langue", valeur = id), ou 
    bien (clÃ© = id, valeur = "code langue").
    
        ParamÃ¨tre :
            <type> : Si 1, alors (clÃ© = "code langue", valeur = id) ;
            si 2, alors (clÃ© = id, valeur = "code langue") ;
            sinon None"""
    
    if type == 1:
        return { "fr-FR" : 1036, "en-US" : 1033 }
    elif type == 2:
        return { 1036 : "fr-FR", 1033 : "en-US" }
    else:
        return None

def lister_les_mots():
    """Renvoie la liste des mots contenus dans le lexique utilisateur.
    
        Remarque : la mÃ©thode GetWords()
            Utilisation des paramÃ¨tres de la mÃ©thode GetWords() (aprÃ¨s tests)
                type de lexique : toujours 1 (lexique utilisateur) sinon
                une erreur se produit.
                GenerationID : inutile, car la derniÃ¨re version du lexique
                est utilisÃ©e.
            La mÃ©thode retourne un tuple (objet ISpeechLexiconWords, generation id).
            Le fonction retourne une liste d'objets de type "mot"."""
    
    lexique = win32com.client.Dispatch("SAPI.SpLexicon")
        
    tuple = lexique.GetWords()
    return tuple[0]

def definir_langue_id(langue):
    """Si la langue est connue, alors retourne l'id correspondant ; et
        sinon renvoie None.
        
        ParamÃ¨tre :
            <langue> : chaÃ®ne du code langue (ex : fr-FR)
                cf http://msdn.microsoft.com/en-us/library/bb813107.aspx"""
    
    dict_langue = dict_code_langue(1)
    if (langue in dict_langue) != True:
        print("Erreur : La langue est inconnue")
        print("    Les langues possibles sont : " + repr(list(dict_langue.keys())))
        return None
    else:
        return dict_langue[langue]
    
def tester_les_phonemes(langue, hta = None, avec_voix = False):
    """Afficher la liste des phonÃ¨mes utilisables par l'objet SpLexicon en fonction d'une langue.
    
        ParamÃ¨tres :
            <avec_voix> : True si les phonÃ¨mes sont prononcÃ©s
            <langue> : chaÃ®ne du code langue (ex : fr-FR)
            <hta> : nom du fichier (sans extension) oÃ¹ la liste 
            est Ã©crite ; l'extension .hta est ajoutÃ©e, le
            fichier est crÃ©e dans le mÃªme dossier, et Ã©crase
            l'ancien"""
    
    # transformer la langue en code
    langue_id = definir_langue_id(langue)
    if langue_id == None:
        return None

    # crÃ©er un objet SpVoice si nÃ©cessaire
    if avec_voix == True:
        voix = win32com.client.Dispatch("SAPI.SpVoice")
    
    print("\n\nLa langue choisie est " + langue)
        
    # crÃ©er le lexique
    lexique = win32com.client.Dispatch("SAPI.SpLexicon")
    
    # crÃ©er la liste des phonÃ¨mes composÃ©s d'un caractÃ¨re unique
    print("CrÃ©ation de la liste de phonÃ¨mes Ã  tester")
    list_phoneme = []
    s = "01234567890&Ã©\"#~'([-|Ã¨`_\Ã§^Ã @)]=+}^Â¨$Â£Â¤Ã¹%*Âµ!Â§:/;.,?<>azertyuiopqsdfghjklmwxcvbnAZERTYUIOPQSDFGHJKLMWXCVBN"
    nb_char = len(s)
    for i in range(0, nb_char):
        list_phoneme.append(s[i])
    
    # ajoute la liste des phonÃ¨mes composÃ©s de 2 caractÃ¨res
    for i in range(0, nb_char):
        char1 = s[i]
        for j in range(0, nb_char):
            char2 = s[j]
            list_phoneme.append(char1 + char2)

    # tester les phonÃ¨mes
    list_phoneme_ok = []
    mot_de_test = "qwsx"
    print("DÃ©but du test")
    nb_phoneme = len(list_phoneme)
    print("Il y a " + repr(nb_phoneme) + " phonÃ¨mes Ã  tester")
    for i in range(0, nb_phoneme): 
        if i % 1000 == 0:
            print(repr(i))
        try: # l'utilisation de certains phonÃ¨mes du franÃ§ais (cf liste) dÃ©clenche une erreur.
            lexique.AddPronunciation(mot_de_test, langue_id, 0, list_phoneme[i]) # Ã  exÃ©cuter 1 fois car conserve
        except:
            pass
        else:
            list_phoneme_ok.append(list_phoneme[i])
            lexique.RemovePronunciation(mot_de_test, langue_id) # attention : erreur si la prononciation n'existe pas
    print("\nFin du test")
    print(repr(len(list_phoneme_ok)) + " phonÃ¨mes sont valides")
            
    # affiche le rÃ©sultat ou crÃ©e le fichier
    nb_phoneme = len(list_phoneme_ok)
    if hta == None:
        print("\nLISTE DES PHONEMES VALIDES\n--------------------------\n")
        for i in range(0, nb_phoneme):
            print(repr(i) + " : " + list_phoneme_ok[i])
            if avec_voix == True:
                lexique.AddPronunciation("windows", 1036, 0, list_phoneme_ok[i])
                voix.Speak("windows")
                lexique.RemovePronunciation("windows", 1036)
                input("<entree> pour quitter le programme")
    else: # crÃ©er un fichier .hta
        list_html = ['<html xmlns="http://www.w3.org/1999/xhtml">', '<head>', 
        '<title>Liste des phonÃ¨mes ' + langue + '</title>',
        '<hta:application', 'id="liste_des_phonemes"', '/>', 
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />', '<style type="text/css">',
        'body { background-color : rgb(255, 239, 213); }', 'h1 { color : rgb(128, 0, 128) }',
        'li { margin-bottom : 2px; font-weight : bold; display : inline}', 
        'input { width : 30px; background-color : rgb(255, 218, 185);}',
        'em { font-weight : bold; text-decoration : underline }',
        '</style>', '<script type="text/javascript">', 'var voix = new ActiveXObject("SAPI.SpVoice");',
        'var lexique = new ActiveXObject("SAPI.SpLexicon");', 'var mot = "' + mot_de_test + '"',
        'function prononcer(phoneme) {',
        'lexique.AddPronunciation(mot,' + str(langue_id) + ', 0, phoneme.value);', 'voix.Speak(mot);',
        'lexique.RemovePronunciation(mot, ' + str(langue_id) + ');', '}', '</script>', '</head>', '<body>',
        '<h1>Liste des phonÃ¨mes dans la langue ' + langue + '</h1>', 
        '<em>Il y a ' + str(nb_phoneme) + ' phoneme(s)</em>',
        '<ol>']
        for i in range(0, nb_phoneme):
            code_html = '<li><input type="button" value="' + list_phoneme_ok[i] + '" onclick="prononcer(this);"/></li>'
            list_html.append(code_html)
        list_html.append('</ol>')
        list_html.append('</body>')
        list_html.append('</html>')
        nom_f = hta + ".hta"
        f = open(nom_f, mode = "w", encoding = "utf_8_sig", errors = "strict")
        for i in range(0, len(list_html)):
            f.write(list_html[i] + "\n")
        f.close()
        print("Le fichier " + nom_f + " est crÃ©e.")
        print("L'application hta doit Ãªtre utilisÃ©e avec la voix correspondante.")
            

            
def ajouter_prononciation(list, langue):
    """Ajouter une liste de prononciations au lexique utilisateur. La liste 
        associe des mots Ã  une prononciation.
    
        ParamÃ¨tres :
            <list> : liste des prononciation
                [i] : mot, [i+1] : prononciation pour i = 0, 1, ...
            <langue> : chaine du code langue

        Remarque :
            1) La mÃ©thode AddPronunciation est utilisÃ©e avec SpeechPartOfSpeech 
            ayant pour valeur SPSUnknown = 0 par dÃ©faut. Cela signifie
            que la prioritÃ© n'est pas gÃ©rÃ©e d'aprÃ¨s le contexte.
            Les prononcations ajoutÃ©es ici sont prÃ©pondÃ©rantes sur le
            lexique de l'application. Si la mÃ©thode est repÃ©tÃ©e pour le mÃªme
            mot, alors la derniÃ¨re prononciation est utilisÃ©e par le moteur
            TTS ou SR, mais le lexique Ã  enregistrÃ© toutes ces prononciations.
            La prononciation s'applique Ã  un mot qui doit Ãªtre sÃ©parÃ© des autres
            mots avec un espace (sÃ©parateur), sinon le moteur TTS ne peut pas
            le reconnaÃ®tre. L'identification des mots tient aussi compte de 
            la casse. Suite aux tests, 
                lexique     texte
                "kHz"       seul le mot "kHz" est reconnu
                "khz"       le mot "khz" est reconnu quelque soit la casse
            Donc, lorsque le moteur TTS recontre un mot (avec la voix Virginie) : 
                1) s'il existe dans le lexique, alors la prononciation est 
                utilisÃ©e. 
                2) Sinon, ce mot est converti en minuscule, puis retour
                Ã  l'Ã©tape 1.
            La suppression du mot Ã  partir du lexique, supprime toutes les
            prononciations.
            2) La prononciation d'un mot "a tendance" Ã  "aspirer" le dernier son
            du mot, celui produit un "effet de fondu" avec la ponctuation ou
            les mot suivants. Pour y remÃ©dier, il suffit d'ajouter le phonÃ¨me 
            '_' en fin de prononciation.
            3) Les mots sont enregistrÃ©s en majuscules (cf doc de la fonction
            analyser_le_texte() dans le module tts.py).
        Important : suite aux tests
            Ne pas inclure de sÃ©parateur (ex : ':') dans un mot du lexique, car
            il n'est pas reconnu.
            Le lexique et la balise <pron> donnent des prononciations Ã©quivalentes.
            Le lexique utilisateur est crÃ©e une seule fois, et il est utilisable
            par toutes les applications. Par contre les balises xml doivent Ãªtre
            insÃ©rÃ©es dans le texte Ã  chaque fois."""

    # transformer la langue en code
    langue_id = definir_langue_id(langue)
    if langue_id == None:
        return False

    # la liste contient un nb pair d'Ã©lÃ©ments
    nb_elt = len(list)
    nb_pair = nb_elt / 2
    if (nb_pair) != int(nb_pair):
        print("Erreur : la liste des prononciations est incomplÃ¨te")
        return False

    lexique = win32com.client.Dispatch("SAPI.SpLexicon")
    for i in range(0, nb_elt, 2):
        mot = list[i].upper() # cf doc analyser_le_texte() - remarque 2)
        try:
            prononciation = list[i + 1] + " _" # cf remarque 2)
            lexique.AddPronunciation(mot, langue_id, SpeechPartOfSpeech.SPSUnknown, prononciation) # cf rem 1)
        except:
            print("Attention : phonÃ¨me incorrect pour le mot " + repr(list[i]))
        else:
            print("Le mot " + mot + " est ajoutÃ©")
    
def supprimer_prononciation(list, langue):
    """Supprimer une liste de prononciations du lexique utilisateur.
    
        ParamÃ¨tres :
            <langue> : chaine du code langue
            <list> : liste des prononciations 
                [i] : mot Ã  supprimer

        Remarque :
            Voir la remarque dans la fonction ajouter_prononciation()."""

    # transformer la langue en code
    langue_id = definir_langue_id(langue)
    if langue_id == None:
        return False

    # supprimer les mots
    lexique = win32com.client.Dispatch("SAPI.SpLexicon")
    for i in range(0, len(list)):
        mot = list[i].upper() # cf doc analyser_le_texte() - remarque 2)
        try:
            lexique.RemovePronunciation(mot, langue_id, SpeechPartOfSpeech.SPSUnknown) # cf rem 1)
        except:
            print("Remarque : le mot " + repr(list[i]) + " n'est pas dans le lexique utilisateur")
        else:
            print("Le mot " + mot + " est supprimÃ©")

def supprimer_le_lexique(langue):
    """Supprime toutes les prononciations du lexique utilisateur pour une langue donnÃ©e.
    
        ParamÃ¨tre :
            <langue> : le code langue
        Renvoie :
            False, si <langue> est erronÃ© ou le lexique est vide"""

    # transformer la langue en code
    langue_id = definir_langue_id(langue)
    if langue_id == None:
        return False

    les_mots = lister_les_mots()
    nb_mots = les_mots.Count
    if nb_mots == 0:
        print("Il n'y a pas de mot Ã  supprimer")
        return False
    
    print("\n\nDÃ©but de la suppression du lexique " + langue)
    lexique = win32com.client.Dispatch("SAPI.SpLexicon")
    SPSUnknown = 0
    for i in range(0, nb_mots):
        mot = les_mots.Item(i)
        if mot.LangId == langue_id:
            lexique.RemovePronunciation(mot.Word, langue_id, SpeechPartOfSpeech.SPSUnknown) # cf rem 1)
            print("Le mot " + mot.Word + " est supprimÃ©")
    print("Fin de la suppression")
    
def afficher_les_mots():
    """Afficher les mots du lexique utilisateur quelque soit la langue."""
            
    les_mots = lister_les_mots()
    nb_mots = les_mots.Count
    print("\n\nIl y a " + repr(nb_mots) + " mot(s) dans le lexique utilisateur")
    if nb_mots == 0:
        return
    
    dict_langue = dict_code_langue(2)
    
    # afficher la liste
    print("\nCONTENU DU LEXIQUE UTILISATEUR\n------------------------------")
    print("Mots".ljust(20) + "Langues".ljust(10) + "Prononciations".ljust(20))
    for i in range(0, nb_mots):
        mot = les_mots.Item(i)
        les_pron = mot.Pronunciations # cf doc fct ajouter_prononciation() : plusieurs prononciations par mot
        langue = ""
        if (mot.LangId in dict_langue) == True:
            langue = dict_langue[mot.LangId]
        pron = les_pron.Item(0).Symbolic # prendre la 1Ã¨re prononciation car elle est active
        print(mot.Word.ljust(20) + langue.ljust(10) + pron.ljust(20))
    print("\n")

# ********************************************************
#       corps principal
# ********************************************************   

if __name__ == "__main__": # le module est exÃ©cutÃ© indÃ©pendament
    pass