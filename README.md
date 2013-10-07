SARAH-Plugin-runstop
====================

### Run, Stop and Restart S.A.R.A.H.


Plugin pour S.A.R.A.H. (http://www.encausse.net) qui permet, vocalement, d'arrêter définitivement ou de redémarrer complètement SARAH.
Il permet également d'utiliser les actions Arrêt/Démarrage/Redémarrage de manière autonome et indépendante de SARAH, directement à partir des scripts.


Egalement :
- Possibilité de paramétrer une phrase d'accueil lors de son démarrage
- Possibilité de lancer Log2Console
- Possibilité de lancer ces modules en mode minimisé ou non
- Possibilité de fermer également Log2Console
- Possibilité de forcer ou non la fermeture de SARAH
- etc...﻿


Bien lire la documentation fournie (index.html) accessible, une fois le plugin installé, via l'interface de SARAH en cliquant sur le Portlet "Run & Stop" !


### ATTENTION
Certaines valeurs du fichier INI (\bin\Config_RunStop.ini) changent selon la version de SARAH !
Les valeurs présentes par défaut sont compatibles jusqu'à la version 3alpha2 (donc 2.9 inclus).
Pensez à bien mettre à jour ces valeurs si besoin est, tout est expliqué dans ce fichier (ouvrez-le avec Notepad par exemple)

#### Dans la section [RUN] :

```VB.net
  ; Version <= 3 alpha 2 : NodeJS=WSRNode.cmd
  ; Version >= 3 beta 1  : NodeJS=Server_NodeJS.cmd
  NodeJS=WSRNode.cmd
  
  ; Version <= 3 alpha 2 : Micro=WSRMicro.cmd
  ; Version >= 3 beta 1  : Micro=Client_Microphone.cmd
  Micro=WSRMicro.cmd
  
  ; Version <= 3 alpha 2 : Kinect=WSRKinect.cmd
  ; Version >= 3 beta 1  : Kinect=Client_Kinect_Audio.cmd OR Kinect=Client_Kinect.cmd
  Kinect=WSRKinect.cmd
  
  ; All versions : Log2Console=Log2Console.exe
  Log2Console=Log2Console.exe
```

#### Dans la section [STOP] :
```VB.net
  ; Version <= 3 alpha 2 : CmdLineNode=WSRNode.cmd
  ; Version >= 3 beta 1  : CmdLineNode=Server_NodeJS.cmd
  CmdLineNode=WSRNode.cmd
  
  ;  All versions : CmdLineConhost=conhost.exe
  CmdLineConhost=conhost.exe
  
  ;  All versions : CmdLineMicro=WSRMacro.exe
  CmdLineMicro=WSRMacro.exe
  
  ;  All versions : CmdLineKinect=WSRMacro_Kinect.exe
  CmdLineKinect=WSRMacro_Kinect.exe
  
  ;  All versions : CmdLineLog2Console=Log2Console.exe
  CmdLineLog2Console=Log2Console.exe
```

