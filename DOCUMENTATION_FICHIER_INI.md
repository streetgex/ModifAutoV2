# Documentation du fichier INI - ModifAuto.Net2010

Cette documentation concerne le projet :

```text
C:\Local\programmation\ModifAuto.Net2010\ModifAuto.Net2010
```

Elle est basee sur la lecture du code source de cette version, notamment
`Form1.vb`, `Gestion.vb`, `PwShell.vb`, `jsonMS.vb`, `ADConfig.vb` et les
classes communes `Commun.vb`, `Json.vb` et `iniFile.vb`.

## Fichier INI utilise

Le programme utilise le fichier INI partage :

```text
\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ScriptStephV2.ini
```

Le chemin est defini dans `Form1.vb` et aussi dans les classes communes.

## Regles generales du fichier INI

Le fichier est organise par sections :

```ini
[SECTION]
parametre=valeur
```

Les lignes commencant par `#` ou `;` sont ignorees lors de la lecture.

Les commentaires en fin de ligne apres un `#` sont retires par le lecteur INI.

Exemple :

```ini
sendMSreport=False # commentaire ignore
```

Les parametres dont le nom se termine par `_encrypted` sont dechiffres
automatiquement.

Exemple :

```ini
AdminScriptPassword_encrypted=valeur_chiffree
```

Ne pas mettre les mots de passe ou jetons en clair.

Attention : le programme ecrit automatiquement `lastExec`. La methode d'ecriture
reecrit le fichier INI complet a partir des sections et cles connues. Il faut
donc eviter les modifications manuelles simultanees pendant une execution.

## Section [GLOBAL]

### SMTPServer

Serveur SMTP utilise pour envoyer les mails.

Format attendu :

```ini
SMTPServer=smtp.exemple.fr
```

Le code tente l'envoi en TLS sur le port 587, puis en cas d'echec sur le port 25
sans TLS.

### AdminScriptLogin

Compte administratif utilise par ModifAuto pour les operations AD, LDAP,
PowerShell et certains appels applicatifs.

Format attendu :

```ini
AdminScriptLogin=userprog
```

### AdminScriptPassword_encrypted

Mot de passe chiffre du compte administratif.

Format attendu :

```ini
AdminScriptPassword_encrypted=valeur_chiffree
```

### MailEquipeReso

Adresse utilisee pour certaines alertes reseau ou destination non associee.

Format attendu :

```ini
MailEquipeReso=equipe-reseau@example.fr
```

Le champ peut contenir plusieurs destinataires separes par `;`.
Les prefixes `Cc:` et `Bcc:` sont supportes par la fonction d'envoi de mail.

Exemple :

```ini
MailEquipeReso=dest1@example.fr;Cc:dest2@example.fr
```

## Section [MODIFAUTO]

### CheminPartage

Chemin reseau utilise pour deposer ou recuperer les fichiers de travail.

Le code utilise notamment les sous-dossiers :

```text
\todo
\cmpttmp
```

Format attendu :

```ini
CheminPartage=\\serveur\partage
```

### dcs

Liste des controleurs de domaine utilisables.

Au demarrage, le programme teste les DC listes et choisit le plus rapide parmi
ceux qui repondent correctement sur LDAP.

Format attendu : FQDN separes par virgules.

```ini
dcs=dc1.example.fr,dc2.example.fr
```

Chaque valeur est ensuite utilisee pour construire des chemins LDAP du type :

```text
LDAP://dc1.example.fr/...
```

### OUUtilisateurs

OU racine contenant les utilisateurs suivis par ModifAuto.

Format attendu : distinguished name Active Directory, sans `LDAP://`.

```ini
OUUtilisateurs=OU=Utilisateurs,DC=example,DC=fr
```

### OUUtilisateursActifs

OU des utilisateurs actifs.

Format attendu :

```ini
OUUtilisateursActifs=OU=Utilisateurs,DC=example,DC=fr
```

Cette valeur peut etre identique a `OUUtilisateurs` si les comptes actifs sont
directement dans l'OU principale.

### OUUtilisateursDesactives

OU des comptes desactives.

Format attendu :

```ini
OUUtilisateursDesactives=OU=Comptes Desactives,OU=Utilisateurs,DC=example,DC=fr
```

### OUUtilisateursExceptions

OU des comptes a traiter comme exceptions.

Format attendu :

```ini
OUUtilisateursExceptions=OU=Exceptions,OU=Utilisateurs,DC=example,DC=fr
```

### OUUtilisateursProvisoires

OU des utilisateurs provisoires.

Format attendu :

```ini
OUUtilisateursProvisoires=OU=Users Provisoires,OU=Utilisateurs,DC=example,DC=fr
```

Le programme met a jour ces comptes, nettoie certaines appartenances et controle
leurs dates d'expiration.

### OUUtilisateursSortis

OU des utilisateurs sortis.

Format attendu :

```ini
OUUtilisateursSortis=OU=Out,OU=Comptes Desactives,OU=Utilisateurs,DC=example,DC=fr
```

Cette OU est utilisee pour les comptes sortis, les exports PST et les traitements
de suppression de boite mail.

### OUUtilisateursExternes

OU des utilisateurs externes.

Format attendu :

```ini
OUUtilisateursExternes=OU=Externes,OU=Utilisateurs,DC=example,DC=fr
```

### OUUtilisateursInvites

OU des utilisateurs invites.

Format attendu :

```ini
OUUtilisateursInvites=OU=Invites,OU=Utilisateurs,DC=example,DC=fr
```

### Group_Admins_du_domaine

Liste autorisee des membres du groupe AD `Admins du domaine`.

Format attendu : `sAMAccountName` separes par virgules.

```ini
Group_Admins_du_domaine=Administrateur,userprog,compteadm
```

Le programme compare cette liste aux membres reels du groupe et peut retirer ou
ajouter des membres pour remettre le groupe en conformite.

### Group_Administrateurs_de_l_entreprise

Liste autorisee des membres du groupe `Administrateurs de l'entreprise`.

Format attendu :

```ini
Group_Administrateurs_de_l_entreprise=Administrateur,userprog
```

### Group_Administrateurs

Liste autorisee des membres du groupe `Administrateurs`.

Format attendu :

```ini
Group_Administrateurs=Administrateur,Administrateurs,Admins du domaine
```

### Group_Administrateurs_du_schema

Liste autorisee des membres du groupe `Administrateurs du schema`.

Format attendu :

```ini
Group_Administrateurs_du_schema=Administrateur,userprog
```

### Group_Administrateurs_DHCP

Liste autorisee des membres du groupe `Administrateurs DHCP`.

Format attendu :

```ini
Group_Administrateurs_DHCP=Administrateur,userprog
```

### EquipesAdministratives

Liste des groupes d'equipes administratives.

Point d'attention important : dans le code actif, cette valeur est lue avec :

```vb
Split(ini.ReadValue("MODIFAUTO", "EquipesAdministratives"))
```

Sans separateur explicite, `Split` separe par espaces. Si la valeur est separee
par des virgules, elle risque donc d'etre lue comme une seule valeur.

Format conforme au code actuel :

```ini
EquipesAdministratives=achat_eq adm-gen_eq finance_eq rh_eq
```

Si une liste avec virgules est conservee, il faut verifier le code ou corriger
le separateur utilise.

### CasExchangeServer

Serveur Exchange utilise pour les sessions PowerShell distantes.

Format attendu :

```ini
CasExchangeServer=serveur-exchange.example.fr
```

Le code construit une URL de ce type :

```text
http://serveur-exchange.example.fr/powershell
```

### DureeMaxJson

Duree maximale toleree pour la creation du fichier JSON.

Format attendu : nombre entier en minutes.

```ini
DureeMaxJson=40
```

Si la creation du fichier depasse cette duree, un mail d'alerte est envoye.

### mailDureeMaxJson

Destinataire du mail d'alerte lorsque la creation du fichier JSON depasse
`DureeMaxJson`.

Format attendu :

```ini
mailDureeMaxJson=dest1@example.fr;Cc:dest2@example.fr
```

Les destinataires multiples sont separes par `;`.
Les prefixes `Cc:` et `Bcc:` sont supportes.

### LoginDirecteur

Login du directeur.

Format attendu :

```ini
LoginDirecteur=login
```

Ce login sert a appliquer des regles particulieres dans la gestion des donnees AD.

### cheminMAJZoneInfo

Chemin de l'executable `MAJZoneInfo.exe` lance en fin de traitement.

Format attendu :

```ini
cheminMAJZoneInfo=C:\Program Files\Script Steph\MAJZoneInfo.exe
```

### dossierArchivePST

Dossier ou sont stockees les archives PST.

Format attendu :

```ini
dossierArchivePST=\\serveur\partage-pst\
```

Valeur par defaut dans le code :

```text
\\Space2\archives-pst\
```

### jourSuppressionPST

Nombre de jours apres lequel les anciens fichiers PST peuvent etre supprimes.

Format attendu : entier.

```ini
jourSuppressionPST=21
```

Valeur par defaut dans le code : `21`.

### dossierPhotos

Dossier reseau contenant les photos RH.

Format attendu :

```ini
dossierPhotos=\\serveur\photos\
```

Valeur par defaut dans le code :

```text
\\Space2\photos-RH\
```

### sendMSreport

Active ou desactive la creation et l'envoi du rapport MicroSesame.

Format attendu : booleen.

```ini
sendMSreport=True
```

ou :

```ini
sendMSreport=False
```

Le traitement est lance uniquement sur certaines plages horaires du matin.

### sendMailOO

Active ou desactive certains mails lies aux officiers orienteurs / MicroSesame.

Format attendu : booleen.

```ini
sendMailOO=False
```

### lastExec

Date et heure de derniere execution de ModifAuto.

Format attendu :

```text
dd/MM/yyyy HH:mm:ss
```

Exemple :

```ini
lastExec=28/04/2026 15:29:43
```

Ce parametre est mis a jour automatiquement en fin d'execution.

## Section [MYIGBMC]

Cette section est utilisee par `Json.GetMyIGBMC()` pour lire certaines donnees
depuis l'API MyIGBMC, notamment les autorisations de diffusion interne et les
photos.

### url

URL de base de l'API MyIGBMC.

Format attendu :

```ini
url=https://my.example.fr/api/
```

### login

Compte API MyIGBMC.

Format attendu :

```ini
login=api_login
```

### password_encrypted

Mot de passe chiffre du compte API.

Format attendu :

```ini
password_encrypted=valeur_chiffree
```

## Section [IGBMCSERVICES]

Cette section est utilisee pour les appels vers l'API IGBMC Services lorsque
`Json.SendJson` est appele avec le systeme par defaut, notamment pour les
personnes, contrats, destinations, emails et listes de diffusion.

### url

URL de base de l'API.

Format attendu :

```ini
url=https://services.example.fr/api/
```

Le code ajoute ensuite l'URI demandee a cette URL.

## Section [MS]

Cette section configure les appels a MicroSesame.

### serveur

Serveur MicroSesame, avec port si necessaire.

Format attendu :

```ini
serveur=serveur-ms.example.fr:443
```

Le code construit ensuite les URL sous la forme :

```text
https://serveur-ms.example.fr:443/api/...
```

### token_encrypted

Jeton API MicroSesame chiffre.

Format attendu :

```ini
token_encrypted=valeur_chiffree
```

Le jeton est envoye dans l'entete HTTP `X-API-KEY`.

## Section [XIVO]

Le code commun contient encore une branche pour Xivo dans `Json.SendJson`.

### XivoRestUser

Utilisateur REST Xivo attendu par le code.

Format attendu :

```ini
XivoRestUser=login_xivo
```

Point d'attention : dans le code actif, la cle lue est `XivoRestUser`.
Si le fichier INI contient seulement `RestUser`, cette branche ne recuperera pas
le login attendu.

### RestPassword_encrypted

Mot de passe REST Xivo chiffre.

Format attendu :

```ini
RestPassword_encrypted=valeur_chiffree
```

### RestServices

URL de base des services REST Xivo.

Format attendu :

```ini
RestServices=https://serveur-xivo:50051/1.1/
```

## Formats utiles

### Booleens

Utiliser `True` ou `False`.

```ini
sendMSreport=True
sendMailOO=False
```

### Listes separees par virgules

Utilisees notamment pour les listes de membres autorises et les DC.

```ini
dcs=dc1.example.fr,dc2.example.fr
Group_Administrateurs_DHCP=Administrateur,userprog
```

### Listes de destinataires mail

Les destinataires sont separes par `;`.

```ini
mailDureeMaxJson=dest1@example.fr;Cc:dest2@example.fr;Bcc:dest3@example.fr
```

### Distinguished names Active Directory

Les OU et groupes doivent etre indiques sous forme DN, sans prefixe `LDAP://`.

```ini
OUUtilisateurs=OU=Utilisateurs,DC=example,DC=fr
```

Le code ajoute lui-meme le prefixe LDAP et le controleur de domaine choisi.

### Parametres chiffres

Les cles suivantes doivent rester chiffrees :

```ini
AdminScriptPassword_encrypted=...
password_encrypted=...
token_encrypted=...
RestPassword_encrypted=...
```

## Parametres presents dans le fichier INI mais non trouves comme lus par cette version

Les parametres suivants peuvent exister dans le fichier INI partage, mais ne sont
pas lus par le code actif compile de `ModifAuto.Net2010` d'apres la recherche
effectuee :

```text
AdminRouter
AdminRouterPassword_encrypted
MailEquipeInfo
MailErrors
OUUtilisateursAdmin
OUUtilisateursAdminInfo
OUautoLinuxAttributs
```

Ils peuvent etre utilises par d'autres programmes ou par d'anciennes versions.

## Points d'attention

- Ne pas stocker de secret en clair.
- Les valeurs `_encrypted` sont dechiffrees automatiquement.
- `lastExec` est modifie automatiquement par le programme.
- Les OU doivent etre des DN Active Directory sans `LDAP://`.
- `dcs` doit contenir des FQDN de controleurs de domaine separes par virgules.
- Les listes `Group_*` sont separees par virgules.
- `EquipesAdministratives` est actuellement separe par espaces dans le code.
- `DureeMaxJson` et `jourSuppressionPST` doivent etre des entiers.
- `sendMSreport` et `sendMailOO` doivent etre des booleens.
- Pour Xivo, le code lit `XivoRestUser`, pas `RestUser`.
