# -*- coding: utf-8 -*-

import lexique
import tts

texte = "Mon nom est Jack La Rancoeur."

#cree la voix
voix_fr = tts.Voix(xml = False)

# parler : kHz n'est pas correctement prononciation
print("Avec erreur de prononciation")
voix_fr.parler(texte)

# ajoute la prononciation du mot "kHz" au lexique utilisateur
prononciations = ["kHz", "k iy l ow & eh r s"]
lexique.ajouter_prononciation(prononciations, "fr-FR")
lexique.lister_les_mots()

# parler : kHz est prononciation correctement
print("Sans erreur de prononciation")
voix_fr.parler(texte)

# supprime le lexique utilisateur en langue française
lexique.supprimer_le_lexique(langue = "fr-FR")
