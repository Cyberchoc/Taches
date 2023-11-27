import os
import datetime

def del_file(chemin_fichier):
    # Spécifiez le chemin complet du fichier que vous souhaitez supprimer
    #chemin_fichier = "/chemin/complet/vers/votre/fichier.txt"

    # Utilisez la fonction os.remove() pour supprimer le fichier
    try:
        os.remove(chemin_fichier)
        print(f"Le fichier {chemin_fichier} a été supprimé avec succès.")
    except OSError as e:
        print(f"Erreur lors de la suppression du fichier {chemin_fichier}: {e}")


def rename_file(ancien_nom_fichier,nouveau_nom_fichier):
    # Spécifiez l'ancien nom de fichier complet avec le chemin
    #ancien_nom_fichier = "/chemin/complet/vers/votre/ancien_fichier.txt"

    # Spécifiez le nouveau nom de fichier complet avec le chemin
    #nouveau_nom_fichier = "/chemin/complet/vers/votre/nouveau_fichier.txt"

    # Utilisez la fonction os.rename() pour renommer le fichier
    try:
        os.rename(ancien_nom_fichier, nouveau_nom_fichier)
        print(f"Le fichier a été renommé de {ancien_nom_fichier} à {nouveau_nom_fichier}.")
    except OSError as e:
        print(f"Erreur lors du renommage du fichier : {e}")

        
def date_du_jour():    
    date_du_jour = datetime.date.today()
    date_formatee = date_du_jour.strftime("%d/%m/%Y")
    #date_formatee = date_formatee.replace("/","_")
    return date_formatee


def heure_du_jour():    
    date_du_jour = datetime.datetime.today()
    date_formatee = date_du_jour.strftime("%H:%M:%S")
    return date_formatee