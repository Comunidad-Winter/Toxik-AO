VERSION 5.00
Begin VB.Form frmHlp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ayuda"
   ClientHeight    =   6225
   ClientLeft      =   2355
   ClientTop       =   1845
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHlp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4296.605
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFAQ 
      Caption         =   "Preguntas &frecuentes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2130
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdBPage 
      Caption         =   "< P�gina &anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdNPage 
      Caption         =   "P�gina &siguiente >"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Ayuda 
      Caption         =   "Introducci�n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Label lblHlp 
         Caption         =   $"frmHlp.frx":0442
         Height          =   4995
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5085
      End
   End
End
Attribute VB_Name = "frmHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EnQuePagina As Integer

Private Sub cmdFAQ_Click()
    lblHlp.Caption = "."
End Sub

Private Sub cmdBPage_Click()
    EnQuePagina = EnQuePagina - 1
    Call CambioPagina
End Sub

Private Sub cmdNPage_Click()
    EnQuePagina = EnQuePagina + 1
    Call CambioPagina
End Sub

Private Sub CambioPagina()

Select Case EnQuePagina

Case 0
    Ayuda.Caption = "Introducci�n"
    lblHlp.Caption = "En ImperiumAO encontrar�s todo un nuevo mundo por explorar, sin fronteras, en el que no hay profesi�n que sobresalga sobre otra y a�n el ambiente se mantiene vivo. A continuaci�n pod�s encontrar algunas de las preguntas m�s frecuentes que se realizan los viajeros de estas tierras."
    cmdBPage.Visible = False
    cmdNPage.Visible = True
Case 1
    Ayuda.Caption = "Empezando a jugar"
    lblHlp.Caption = "En un principio, tendr�s, como m�nimo, ropa, agua, comida, un mapa y un arma. Un buen lugar para empezar es el Newbie Dungeon, que posee dos zonas. En la primera encontrar�s tenebrosas criaturas que te desafiar�n, en la segunda una zona en donde entrenar tus habilidades de combate. Estando un alguna de las ciudades iniciales, podr�s tomar el teleport al mismo f�cilmente. Mientras subas de nivel, ir�s ganando oro, con lo que puedes comprar hechizos, armas, ropa, etc. Cuando tu nivel sea 13 o superior, dejar�s de ser newbie y ya no ganar�s oro al subir de nivel. Deber�s ir pensando en acceder a los bosques, los dungeons, y todo el mundo que te espera."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 2
    Ayuda.Caption = "Entrenando"
    lblHlp.Caption = "Cada clase es �nica. Al haber elegido cierta clase, estas sujeto a ser mejor en algunas �reas que en otras. Aprov�chalas. Si eres mago, ded�cate a la magia. Si eres guerrero, ded�cate al combate con armas. Cuando subas de nivel, ganar�s skillpoints, los cuales puedes decidir utilizar o no. Lo m�s conveniente es dejarlos para cuando seas de mayor nivel. Por ahora, mientras entrenes, ir�s ganando naturalmente los skillpoints. A mayor skill en tal �rea, mejor es el rendimiento que tendr� tu personaje al realizar la actividad. Explora: es la mejor manera de conocer el mundo en el que vive tu personaje. Encontrar�s miles de criaturas, nuevos retos y gente con quien afrontarlos. Los mejores lugares para comenzar a entrenar luego de haber dejado el Newbie Dungeon son el Bosque Dorck (Mapas 39 y 38) y las Colinas de Nix (Mapa 35 y alrededores) En estas zonas encontrar�s tus primeros retos: Ara�as Gigantes y Zombies."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 3
    Ayuda.Caption = "Interactuando"
    lblHlp.Caption = "A lo largo de tus viajes, encontrar�s muchas maneras de interactuar con el mundo en el que tu personaje vive. Encontrar�s ciudades, dungeons, bosques, y muchos lugares que llamar�n tu atenci�n. Si est�s en una ciudad una buena idea es, si es que �sta tiene un puerto o r�o cercano, es pescar. Puedes adquirir una ca�a o red de pesca en el gremio de pescadores. Otra opci�n es talar. A medida que explores las diferentes maneras de ganarte la vida, decidir�s cu�l crees como m�s apropiada. No dejes de probar la miner�a y herrer�a, pues es una actividad muy productiva e �til si es que tu personaje tiene habilidades de combate. En ImperiumAO estas �reas han sido rebalanceadas para obtener, as�, un mundo m�s justo y pr�spero."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 4
    Ayuda.Caption = "Facci�nes"
    lblHlp.Caption = "En la medida de que llegues a niveles superiores, podr�s decidir enlistarte o no en una facci�n. Puedes elegir las Tropas Reales o bien las Legiones del Caos. Para ingresar en las Tropas Reales debes haber matado, como m�nimo, cincuenta criminales. Si en alg�n momento asesinaste gente inocente no ser�s admitido. Tu nivel no deber� ser inferior a 20. Por otro lado, si tu deseo es ingresar en las Legiones del Caos, deber�s haber matado como m�nimo 125 ciudadanos y ser nivel 25 o superior. Cuando llegues a la centena de criminales o ciudadanos asesinados, ser�s recompensado y, adem�s de ganar algo de experiencia, subir�s un rango de jerarqu�a, lo que, adem�s de posibilitarte nuevos �tems especiales, har� que tu rango sea superior. Recuerda, no hay vuelta atr�s a la hora de seleccionar una facci�n. Se cuidadoso a la hora de elegir."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 5
    Ayuda.Caption = "Magia"
    lblHlp.Caption = "ImperiumAO fue creado ni mas ni menos que con la quinta esencia, lo que hace a la magia y a lo desconocido. Aprov�chala. Est� en todas partes. Pues, entonces, se han desarrollado nuevos tipos de hechizos. Los hechizos de metamorfosis permiten que el cuerpo de tu personaje cambie a determinada forma, temporalmente. Esto no s�lo te da ventajas t�cnicas sino tambi�n de combate. Si aparentas ser m�s poderoso que una criatura, esta simplemente te ignorar�. Puedes conseguir estos hechizos por precios razonables en la ciudad de Nueva Esperanza. Por otro lado, la magia se ha expandido tanto, que hay no s�lo hechizos de metamorfosis, sino tambi�n de materializaci�n, y mucho m�s que tienes por explorar. Se dice que hay magos con poderes tales, que pueden abrir portales planares con su magia para viajar entre el espacio. Muchos otros dicen haber visto magia con poder suficiente para matar hasta al m�s poderoso guerrero de estas tierras."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 6
    Ayuda.Caption = "Viajando"
    lblHlp.Caption = "Hay, pues, varias maneras no convencionales de viajar hacia otras tierras. Si hay un puerto en las cercan�as y quieres llegar r�pido a tu destino �Por qu� no tomar un barco? Es muy f�cil, s�lo debes acercarte el muelle, comerciar con el pirata y luego mostrarle el pasaje, as�, llegar�s al destino elegido. Los pases no caen al morir y no pueden ser robados. Lleva un mapa. Te salvar� un situaci�nes comprometedoras. Recuerda que para viajar a zonas lejanas debes poder usar una Barca."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 7
    Ayuda.Caption = "Objetos m�gicos"
    lblHlp.Caption = "Anillos, brazaletes, botas, armas, armaduras, escudos y todo tipo de extra�os objetos han sido impregnados de magia por los m�s poderosos y Arcanos hechizeros. Ellos, pues, han dispersado esos codiciados �tems por todas estas tierras. Actualmente en poder de las criaturas m�s poderosas, estos objetos pueden darte propiedades m�gicas que van desde la inmunidad a hechizos, hasta el aumento del tiempo en que recuperas tus puntos de mana. Al morir una de estas criaturas, si tienes suerte, podr�s recibir uno de estos objetos."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 8
    Ayuda.Caption = "Cazando Recompensas"
    lblHlp.Caption = "En la ciudad de Rinkel no s�lo se han establecido piratas, sino todo un mercado negro organizado, que est� constantemente trabajando por la gente que necesita, por ejemplo, que alguien sea asesinado a toda costa. Revisa la ciudad, all� encontrar�s un cazarecompensas con el que podr�s, adem�s de ofrecer dinero por la cabeza de un usuarios, ver qui�nes son buscados y as� poder ganar dinero de una manera f�cil y efectiva."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 9
    Ayuda.Caption = "Familiares y mascotas permanentes"
    lblHlp.Caption = "Los familiares son criaturas de otros planos que, gracias a tu poder m�gico, han decidido permancer junto a ti. Durante la creaci�n de tu personaje, si eres Mago, podr�s elegir un Familiar. Por otro lado, si eres Cazador o Druida, tendr�s la posibilidad de tener una mascota permanente. Los familiares y mascotas permanentes tienen niveles de vida, mueren, revivien, se hacen m�s fuertes entrenando, exactamente como lo haces vos. Una de estas criaturas ser� tu fiel seguidor y te ayudar� en los momentos m�s dif�ciles. Si tienes un familiar, podr�s ver sus estad�sticas al ver las tuyas, haciendo click en el Boton que se encuentra en el men� inferior. Todos ganan habilidades especiales al llegar a diferentes niveles."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 10
    Ayuda.Caption = "Nuevas fronteras"
    lblHlp.Caption = "Nuevos �mbitos se han descubierto. Si eres un Bardo o tu conocimiento en Artes Marciales es 10 o superior, al combatir sin armas, tendr�s una probabilidad (que depende de tu skill) de hacer movimientos especiales al atacar. Estos movimientos terminan en efectos que van desde par�lisis, ceguera, estupidez, hasta el desarmado de la v�ctima (la v�ctima perder� su arma) Al recibir da�o con hechizos, tu personaje ir� aprendiendo diferentes modos de resisitirlo. A mayor resistencia, menor da�o te har�n los hechizos. Shurikens, lanzas, jabalinas, dagas arrojadizas, y much�simas otras armas te esperan. Con tener s�lo las municiones alcanza, no necesitas nada m�s."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 11
    Ayuda.Caption = "Macros Integrados"
    lblHlp.Caption = "Se han integrado en el cliente una serie de teclas para mejorar la din�mica del juego. Su uso es completamente legal. Est� TOTALMENTE prohibido el uso de cualquier programa externo al juego. Para poder utilizar estas teclas, prueba presionando F1, F2, F3, etc..."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 12
    Ayuda.Caption = "Game Masters"
    lblHlp.Caption = "Este mundo se encuentra regido por dioses, quienes mantienen el �rden y est�n constantemente trabajando por su mejor�a. Puedes invocarlos escribiendo /GM. De este modo, los dioses acudir�n en tu ayuda. Los mismos son f�cilmente reconocibles, su nombre, habla, y descripci�n se encuentra en verde."
    cmdBPage.Visible = True
    cmdNPage.Visible = False
End Select
End Sub
