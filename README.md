# Email Automation with Excel VBA

## Objectif
Ce projet reproduit un cas simple d'automatisation métier : envoyer plusieurs emails avec pièces jointes depuis un fichier Excel via Outlook.

## Contexte
Ce projet a été reconstruit à partir d'un cas concret d'automatisation d'une tâche répétitive : l'envoi d'emails avec pièces jointes à plusieurs destinataires.

L'objectif est de démontrer une première logique d'automatisation appliquée à un besoin métier simple.

## Fonctionnalités
- lecture des destinataires depuis Excel
- gestion des champs To / CC / BCC
- objet et message personnalisés
- ajout de pièces jointes
- envoi d'emails via Outlook
- mise à jour du statut d'envoi
- macros utilitaires pour effacer certaines colonnes
- sélection de fichiers pour les pièces jointes

## Technologies utilisées
- Excel
- VBA
- Outlook Desktop
- Git / GitHub

## Structure du projet
- `Envoi_Mails_VBA.xlsm` : fichier Excel principal
- `Module_Envoi_Mails.bas` : module VBA exporté
- `docs/` : documents PDF fictifs de démonstration
- `screenshots/` : captures de l'interface

## Captures
### Interface Excel
![Interface Excel](screenshots/excel-interface.png)

### Résultat dans Outlook
![Mail reçu dans Outlook](screenshots/outlook-mail-recu.png)

## Limites actuelles
- dépendance à Outlook classique
- gestion d'erreur simple
- interface volontairement basique

## Améliorations possibles
- journalisation des erreurs
- interface plus ergonomique
- gestion avancée des erreurs
- adaptation future en Python ou via API

## Auteur
Projet réalisé par Cindy K / Nereais