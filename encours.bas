Attribute VB_Name = "Module1"
'Encours

'Déclaration des variables publiques
Public RSH As Worksheet     'Feuille récap
Public TASH As Worksheet    'Feuille tableaux analytiques
Public TPSH As Worksheet    'Feuille tableau analytique prêts
Public ASH As Worksheet     'Feuille encours moyen actions
Public TSH As Worksheet     'Feuille encours moyen taux
Public YSH As Worksheet     'Feuille encours moyen transactions
Public PSH As Worksheet     'Feuille encours moyen prêts
Public SSH As Worksheet     'Feuille soldes
Public BSH As Worksheet     'Feuille base de données
Public BPSH As Worksheet    'Feuille base de données prêts
Public ESH As Worksheet     'Feuille échéancier
'Public XSH As Worksheet     'Feuille de tests
'Autre
Public di As Integer        'Date de première opération
Public dd As String         'Date de la dernière opération
Public s As New Collection  '
Public c As New Collection  'Collection de titres pour les tableaux analytiques
Public arr() As Variant

'Portefeuilles titre -> Numéros de comptes correspondants
Public Const P_OBLIG_INVEST As String = "411100 411600 412100 412200 412300 412400"
Public Const P_OBLIG_PLACT As String = "311100 313100 313200 313300 313400"
Public Const P_ACTION_PART As String = "422100 422200 422300 422500 423100 423200 423500"
Public Const P_ACTION_PLACT As String = "315100 315200"
Public Const P_ACTION_TRANS As String = "305200"
'Public Const P_ACTION_TRANS As Variant = Array("305200")

'Index
'Feuille 'Données Titres' (BSH)
Public Const BSHC As String = "AK"  'Dernière colone
Public Const IBC As Integer = 5     'Numéro de compte
Public Const IBD As Integer = 11     'Date comptable
Public Const IBV As Integer = 37     'Valeur comptable
Public Const IBCC As Integer = 22    'Code du titre
Public Const IBP As Integer = 20     'Poste (achat, vente, reclassement, augmentation/réduction capital)
Public Const IBO As Integer = 4      'op_evenement (pour le cas du portefeuille TRANS. Les opérations ne doivent pas être prises en comptes si champ non vide)
'Feuille 'Données prêts' (BPSH)
Public Const IPBO As Integer = 1    'Numéro d'opération
Public Const IPBP As Integer = 2    'Poste
Public Const IPBC As Integer = 3    'Contrepartie
Public Const IPBD As Integer = 4    'Date de valeur
Public Const IPBR As Integer = 5    'Date de remboursement
Public Const IPBN As Integer = 6    'Nominal
Public Const IPBV As Integer = 12    'Montant remboursement
'public ipbl as Integer      'PNL
'Public ipbc As Integer      'Numéro de compte
'Public ipbd As Integer      'Date comptable
'Public ipbv As Integer      'Valeur comptable
'Feuille 'Echéancier' (ESH)
Public Const IEN As Integer = 1     'Colonne numéro d'opération
Public Const IED As Integer = 2     'Colone date
Public Const IET As Integer = 3     'Colone tombée
'Feuille 'soldes' (SSH)
Public Const ISL As Integer = 6         'Ligne de départ du tableau des soldes
Public Const ISLAT As Integer = ISL + 1 'Ligne du solde actions transactions
Public Const ISLAP As Integer = ISL + 2 'Ligne du solde actions placement
Public Const ISLAI As Integer = ISL + 3 'Ligne du solde actions participation
Public Const ISLOP As Integer = ISL + 4 'Ligne du solde obligations placement
Public Const ISLOI As Integer = ISL + 5 'Ligne du solde obligations investissement
Public Const ISLP As Integer = ISL + 6  'Ligne du solde prêts
Public Const ISC As Integer = 5         'Colonne de départ du tableau des soldes
Public Const ISAL As Integer = 16    'Ligne de départ du tableau des soldes actions
'Public Const isac As Integer = 5    'Colonne de départ du tableau des soldes actions
Public isac As Integer    'Colonne de départ du tableau des soldes actions
Public Const ISACN As Integer = 2   'Colonne nom
Public Const ISACC As Integer = 3   'Colonne code
Public Const ISACP As Integer = 4   'Colonne portefeuille
'Feuille 'encours'
Public Const IEL As Integer = 5     'Ligne de début du premier tableau
Public Const IEC As Integer = 2     'Colonne de début du premier tableau
Public Const IECE As Integer = 8    'Espace entre les deux tableaux
Public Const IECM As Integer = IEC + 1    'Colonne mouvement du premier tableau
Public Const IECA As Integer = IEC + 2  'Colonne acquisition du premier tableau
Public Const IECC As Integer = IEC + 3  'Colonne cession du premier tableau
Public Const IECS As Integer = IEC + 4  'Colonne solde du premier tableau
Public Const IECD As Integer = IEC + 5  'Colonne durèe du premier tableau
Public Const IECP As Integer = IEC + 6  'Colonne des pondérations (solde * durée) du premier tableau
Public Const IEC2 As Integer = IEC + IECE  'Colonne de début de la 2ème série de tableaux
Public Const IEC2M As Integer = IEC2 + 1  'Colonne mouvement du 2ème tableau
Public Const IEC2A As Integer = IEC2 + 2 'Colonne acquisition du 2ème tableau
Public Const IEC2C As Integer = IEC2 + 3 'Colonne cession du 2ème tableau
Public Const IEC2S As Integer = IEC2 + 4 'Colonne solde du 2ème tableau
Public Const IEC2D As Integer = IEC2 + 5 'Colonne durèe du 2ème tableau
Public Const IEC2P As Integer = IEC2 + 6 'Colonne des pondérations (solde * durée) du 2ème tableau
'Feuille 'Récap' (RSH)
Public Const IRC As Integer = 3         'Colonne de début des soldes comptables
Public Const IRPDD As String = "H3"     'Case date de début (prêts)
Public Const IRPDF As String = "H4"     'Case date de fin (prêts)
Public Const IRATE As Integer = 9       'Ligne des encours actions transactions
Public Const IRAPE As Integer = 10      'Ligne des encours actions placement
Public Const IRAIE As Integer = 11      'Ligne des encours actions participation
Public Const IROPE As Integer = 12      'Ligne des encours taux placement
Public Const IROIE As Integer = 13      'Ligne des encours taux investissement
Public Const IRPE As Integer = 14       'Ligne des encours prêts
Public Const IRAT As Integer = 18       'Ligne du solde actions transactions
Public Const IRAP As Integer = 19       'Ligne du solde actions placement
Public Const IRAI As Integer = 20       'Ligne du solde actions participation
Public Const IROP As Integer = 21       'Ligne du solde taux placement
Public Const IROI As Integer = 22       'Ligne du solde taux investissement
Public Const IRP As Integer = 23        'Ligne du solde prêts
Public Const IRRE As String = "D9:O13"      'Espace à réinitialiser (encours moyens)
Public Const IRRS As String = "D18:O22"     'Espace à réinitialiser (soldes comptables)
Public Const IRREP As String = "D14:O14"    'Espace à réinitialiser (encours moyens prêts)
Public Const IRRSP As String = "D23:O23"    'Espace à réinitialiser (soldes comptables prêts)

Sub variables()

Set RSH = ThisWorkbook.Worksheets("Récap")
Set TASH = ThisWorkbook.Worksheets("Tableaux Analytiques Titres")
Set TPSH = ThisWorkbook.Worksheets("Tableau Analytique Prêts")
Set ASH = ThisWorkbook.Worksheets("EM Actions")
Set TSH = ThisWorkbook.Worksheets("EM Oblig")
Set YSH = ThisWorkbook.Worksheets("EM Trans")
Set PSH = ThisWorkbook.Worksheets("EM Prêts")
Set SSH = ThisWorkbook.Worksheets("Soldes Début")
Set BSH = ThisWorkbook.Worksheets("Données Titres")
Set BPSH = ThisWorkbook.Worksheets("Données Prêts")
Set ESH = ThisWorkbook.Worksheets("Echéancier Prêts")
'Set xsh = ThisWorkbook.Worksheets("test")

'Emplacements dans la feuille 'Soldes' partie soldes par titre
'Utiliser avec ssh.Cells(isal, isac)
isac = 5        'Colonne de départ

End Sub

Sub commencer()

'TESTDateDebut = Now

optimisationDébut
variables
réinitialiser

'Encours titres
encours_titres                  'Encours par opération/mois
'encours_titres_2                'Encours par opération/mois (array) (plus rapide)
init_tableaux_analytiques


optimisationFin

'MsgBox DateDiff("s", TESTDateDebut, Now)

End Sub
Sub commencer_prets()

'TESTDateDebut = Now

optimisationDébut
variables
réinitialiser_prets

'Encours prêts
encours_prets_2
init_tableau_analytique_prets

optimisationFin

'MsgBox DateDiff("s", TESTDateDebut, Now)

End Sub

Sub encours_titres()

'Filtre et ordonne la base de données par date comptable du plus ancien au plus récent
trier_par_dates BSH, "K"

'Détermine l'année de base
Dim y As Integer
y = Year(BSH.Cells(2, IBD))
di = Year(BSH.Cells(2, IBD))

'Crée le contenant des calculs
Set s = New Collection
s.Add New Collection, "Action"
s.Add New Collection, "Tx"

s.Item("Action").Add New Collection, "TRANS"
s.Item("Action").Add New Collection, "PLACT"
s.Item("Action").Add New Collection, "PART"

s.Item("Tx").Add New Collection, "PLACT"
s.Item("Tx").Add New Collection, "INVEST"

For Each i In s
    For Each j In i
        j.Add CDate("01/01/" & y), "date"       'Date
        j.Add 0, "Mvt"          'Mouvement
        j.Add 0, "Pond."        'Solde * Durée
    Next j
Next i

'Collection tableaux analytiques
Set c = New Collection
c.Add New Collection, "TRANS"
c.Add New Collection, "PART"
c.Add New Collection, "PLACT"

'Cherche le solde de départ correspondant à l'année de base dans le tableau des soldes
With SSH
i = ISC
Do Until IsEmpty(.Cells(ISL, i))
    If y - 1 = .Cells(ISL, i) Then
        s.Item("Action").Item("TRANS").Add .Cells(ISLAT, i), "solde"
        s.Item("Action").Item("PLACT").Add .Cells(ISLAP, i), "solde"
        s.Item("Action").Item("PART").Add .Cells(ISLAI, i), "solde"
        s.Item("Tx").Item("PLACT").Add .Cells(ISLOP, i), "solde"
        s.Item("Tx").Item("INVEST").Add .Cells(ISLOI, i), "solde"
    End If
    i = i + 1
Loop
End With

'Cherche les soldes de départ des tableaux analytiques
With SSH
'Cherche la colonne correspondante à l'année de base dans le tableau des soldes actions
i = isac
Do Until IsEmpty(.Cells(ISAL, i))
    If y - 1 = .Cells(ISAL, i) Then isac = i
    i = i + 1
Loop
'Compile la liste des actions par portefeuilles
i = ISAL + 1
Do Until IsEmpty(.Cells(i, ISACN))
    For Each j In Array("TRANS", "PLACT", "PART")
        If j = .Cells(i, ISACP) Then
            c.Item(.Cells(i, ISACP)).Add New Collection, .Cells(i, ISACC)
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add .Cells(i, ISACC), "code"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add .Cells(i, ISACN), "nom"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add .Cells(i, isac), "solde"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add .Cells(i, isac), "solde fin"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add 0, "acquisition"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add 0, "augmentation capital"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add 0, "reclassement"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add 0, "réduction capital"
            c.Item(.Cells(i, ISACP)).Item(.Cells(i, ISACC)).Add 0, "cession"
        End If
    Next j
    i = i + 1
Loop
End With

'Calcule le solde comptable
With BSH
i = 2
jat = 0     'Compteur de lignes pour encours actions transactions
jap = 0     'Compteur de lignes pour encours actions placement
jai = 0     'Compteur de lignes pour encours actions participation
jtp = 0     'Compteur de lignes pour encours taux placement
jti = 0     'Compteur de lignes pour encours taux investissement
dd = .Cells(i, IBD)      'Date de la dernière opération
affiche_titres_encours_op YSH, IEC, "TRANS"
affiche_titres_encours_op ASH, IEC2, "PLACT"
affiche_titres_encours_op ASH, IEC, "PART"
affiche_titres_encours_op TSH, IEC2, "PLACT"
affiche_titres_encours_op TSH, IEC, "INVEST"
Do Until IsEmpty(.Cells(i, IBD))

    If Month(dd) = Month(.Cells(i, IBD)) - 1 Then    'Solde mensuel
        calcul_mensuel "Action", "TRANS", IRAT, IRATE, i
        calcul_mensuel "Action", "PLACT", IRAP, IRAPE, i
        calcul_mensuel "Action", "PART", IRAI, IRAIE, i
        calcul_mensuel "Tx", "PLACT", IROP, IROPE, i
        calcul_mensuel "Tx", "INVEST", IROI, IROIE, i
    End If
    
    'Détermine la date de la dernière opération
    If CDate(.Cells(i, IBD)) > CDate(dd) Then
        dd = .Cells(i, IBD)
    End If
    
    'Action TRANS
    'If .Cells(i, IBC) = P_ACTION_TRANS Then
    For Each j In Split(P_ACTION_TRANS)
        If .Cells(i, IBC) = j Then
            d = s.Item("Action").Item("TRANS").Item("date")
            If d <> .Cells(i, IBD) Then
                jat = jat + 1
            End If
            calculs "Action", "TRANS", YSH, i, jat - 1, IEC, IECS, IECD, IECP, IECM, IECA, IECC, IRAT, IRATE
            If IsEmpty(.Cells(i, IBO)) Then calculs_tableaux_analytiques "TRANS", i     'Pour éviter de prendre en compte les écritures (ext ATPD et équivalent)
        End If
    'End If
    Next j
    
    'Action PLACT
    For Each j In Split(P_ACTION_PLACT)
        If .Cells(i, IBC) = j Then
            d = s.Item("Action").Item("PLACT").Item("date")
            If d <> .Cells(i, IBD) Then
                jap = jap + 1
            End If
            calculs "Action", "PLACT", ASH, i, jap - 1, IEC2, IEC2S, IEC2D, IEC2P, IEC2M, IEC2A, IEC2C, IRAP, IRAPE
            calculs_tableaux_analytiques "PLACT", i
        End If
    Next j
    
    'Action PART
    For Each j In Split(P_ACTION_PART)
        If .Cells(i, IBC) = j Then
            d = s.Item("Action").Item("PART").Item("date")
            If d <> .Cells(i, IBD) Then
                jai = jai + 1
            End If
            calculs "Action", "PART", ASH, i, jai - 1, IEC, IECS, IECD, IECP, IECM, IECA, IECC, IRAI, IRAIE
            calculs_tableaux_analytiques "PART", i
        End If
    Next j

    'OBL PLACT
    For Each j In Split(P_OBLIG_PLACT)
        If .Cells(i, IBC) = j Then
            d = s.Item("Tx").Item("PLACT").Item("date")
            If d <> .Cells(i, IBD) Then
                jtp = jtp + 1
            End If
            calculs "Tx", "PLACT", TSH, i, jtp - 1, IEC2, IEC2S, IEC2D, IEC2P, IEC2M, IEC2A, IEC2C, IROP, IROPE
        End If
    Next j
    
    'OBL INVEST
    For Each j In Split(P_OBLIG_INVEST)
        If .Cells(i, IBC) = j Then
            d = s.Item("Tx").Item("INVEST").Item("date")
            If d <> .Cells(i, IBD) Then
                jti = jti + 1
            End If
            calculs "Tx", "INVEST", TSH, i, jti - 1, IEC, IECS, IECD, IECP, IECM, IECA, IECC, IROI, IROIE
        End If
    Next j
    
    i = i + 1
Loop
End With

''
correction_fin "Action", "TRANS", YSH, jat, IEC, IECS, IECD, IECP, IECM, IECA, IECC, IRAT, IRATE
correction_fin "Action", "PLACT", ASH, jap, IEC2, IEC2S, IEC2D, IEC2P, IEC2M, IEC2A, IEC2C, IRAP, IRAPE
correction_fin "Action", "PART", ASH, jai, IEC, IECS, IECD, IECP, IECM, IECA, IECC, IRAI, IRAIE
correction_fin "Tx", "PLACT", TSH, jtp, IEC2, IEC2S, IEC2D, IEC2P, IEC2M, IEC2A, IEC2C, IROP, IROPE
correction_fin "Tx", "INVEST", TSH, jti, IEC, IECS, IECD, IECP, IECM, IECA, IECC, IROI, IROIE

colorer_tableau YSH, IEL, IEC, IECP
colorer_tableau ASH, IEL, IEC2, IEC2P
colorer_tableau ASH, IEL, IEC, IECP
colorer_tableau TSH, IEL, IEC2, IEC2P
colorer_tableau TSH, IEL, IEC, IECP

End Sub
Sub affiche_titres_encours_op(sh As Worksheet, col As Integer, portefeuille As String)

With sh
    
    'titres = "Date,Mvt,Acquisition,Cession,Solde,Durée,Solde * durée"
    'If portefeuille = "PRETS" Then titres = "Date,Mvt,Octroi,Remboursement,Solde,Durée,Solde * durée"
    titres = Array("Date", "Mvt", "Acquisition", "Cession", "Solde", "Durée", "Solde * durée")
    If portefeuille = "PRETS" Then titres = Array("Date", "Mvt", "Octroi", "Remboursement", "Solde", "Durée", "Solde * durée")
    
    'Grand titre
    lignetitre = IEL - 2
    .Cells(lignetitre, col) = portefeuille
        
    'Sous-titres (array)
    lignetitre = IEL - 1
    .Range(.Cells(lignetitre, col), .Cells(lignetitre, UBound(titres) + col)) = titres
    
End With

End Sub


Function calcul_mensuel(item1, item2, rl, rle, i)
    
    RSH.Cells(rl, IRC + Month(dd)) = s.Item(item1).Item(item2).Item("solde")
    'Encours mensuel
    d = s.Item(item1).Item(item2).Item("date")
    dn = DateSerial(Year(dd), Month(dd) + 1, 0) - CDate(d) + 1      'dernière opération du mois -> fin de mois + 1
    j = s.Item(item1).Item(item2).Item("Pond.") + s.Item(item1).Item(item2).Item("solde") * dn
    RSH.Cells(rle, IRC + Month(dd)) = j / _
    (DateSerial(Year(dd), Month(dd) + 1, 0) - _
    (CDate("01/01/" & Year(BSH.Cells(2, IBD))) - 1))
    
End Function
Function calculs(item1, item2, sh, i, ligne, c, cs, cd, cp, cm, ca, cc, rl, rle)

With sh
    ligne2 = IEL + ligne
    d = s.Item(item1).Item(item2).Item("date")
    If d <> BSH.Cells(i, IBD) Then
        .Cells(ligne2, c) = s.Item(item1).Item(item2).Item("date")
        .Cells(ligne2, cs) = s.Item(item1).Item(item2).Item("solde")
        .Cells(ligne2, cs).NumberFormat = "#,##0.00"
        du = CDate(BSH.Cells(i, IBD)) - CDate(d)
        .Cells(ligne2, cd) = du
        .Cells(ligne2, cp) = du * s.Item(item1).Item(item2).Item("solde")
        .Cells(ligne2, cp).NumberFormat = "#,##0.00"
        'Mvt
        .Cells(ligne2, cm) = s.Item(item1).Item(item2).Item("Mvt")
        .Cells(ligne2, cm).NumberFormat = "#,##0.00"
        If s.Item(item1).Item(item2).Item("Mvt") > 0 Then
            .Cells(ligne2, cc) = s.Item(item1).Item(item2).Item("Mvt")
            .Cells(ligne2, cc).NumberFormat = "#,##0.00"
        Else
            .Cells(ligne2, ca) = -s.Item(item1).Item(item2).Item("Mvt")
            .Cells(ligne2, ca).NumberFormat = "#,##0.00"
        End If
        s.Item(item1).Item(item2).Remove "Mvt"
        s.Item(item1).Item(item2).Add BSH.Cells(i, IBV), "Mvt"
        'Somme: Solde * Durée
        j = s.Item(item1).Item(item2).Item("Pond.") + du * s.Item(item1).Item(item2).Item("solde")
        s.Item(item1).Item(item2).Remove "Pond."
        s.Item(item1).Item(item2).Add j, "Pond."
    Else
        j = s.Item(item1).Item(item2).Item("Mvt") + BSH.Cells(i, IBV)
        s.Item(item1).Item(item2).Remove "Mvt"
        s.Item(item1).Item(item2).Add j, "Mvt"
    End If
    j = s.Item(item1).Item(item2).Item("solde") - BSH.Cells(i, IBV)
    s.Item(item1).Item(item2).Remove "solde"
    s.Item(item1).Item(item2).Add j, "solde"
    s.Item(item1).Item(item2).Remove "date"
    s.Item(item1).Item(item2).Add BSH.Cells(i, IBD), "date"
End With

End Function

Function calculs_2(item1, item2, sh, i, ligne, c, cs, cd, cp, cm, ca, cc, rl, rle)

With sh
    ligne2 = IEL + ligne
    d = s.Item(item1).Item(item2).Item("date")
    If d <> arr(i, IBD) Then
        '.Cells(ligne2, c) = s.Item(item1).Item(item2).Item("date")
        .Cells(ligne2, c) = Format(s.Item(item1).Item(item2).Item("date"), "mm/dd/yyyy")
        .Cells(ligne2, cs) = s.Item(item1).Item(item2).Item("solde")
        .Cells(ligne2, cs).NumberFormat = "#,##0.00"
        du = CDate(arr(i, IBD)) - CDate(d)
        .Cells(ligne2, cd) = du
        .Cells(ligne2, cp) = du * s.Item(item1).Item(item2).Item("solde")
        .Cells(ligne2, cp).NumberFormat = "#,##0.00"
        .Cells(ligne2, cm) = s.Item(item1).Item(item2).Item("Mvt")
        .Cells(ligne2, cm).NumberFormat = "#,##0.00"
        If s.Item(item1).Item(item2).Item("Mvt") > 0 Then
            .Cells(ligne2, cc) = s.Item(item1).Item(item2).Item("Mvt")
            .Cells(ligne2, cc).NumberFormat = "#,##0.00"
        Else
            .Cells(ligne2, ca) = -s.Item(item1).Item(item2).Item("Mvt")
            .Cells(ligne2, ca).NumberFormat = "#,##0.00"
        End If
        s.Item(item1).Item(item2).Remove "Mvt"
        s.Item(item1).Item(item2).Add arr(i, IBV), "Mvt"
        'Somme: Solde * Durée
        j = s.Item(item1).Item(item2).Item("Pond.") + du * s.Item(item1).Item(item2).Item("solde")
        s.Item(item1).Item(item2).Remove "Pond."
        s.Item(item1).Item(item2).Add j, "Pond."
    Else
        j = s.Item(item1).Item(item2).Item("Mvt") + arr(i, IBV)
        s.Item(item1).Item(item2).Remove "Mvt"
        s.Item(item1).Item(item2).Add j, "Mvt"
    End If
    j = s.Item(item1).Item(item2).Item("solde") - arr(i, IBV)
    s.Item(item1).Item(item2).Remove "solde"
    s.Item(item1).Item(item2).Add j, "solde"
    s.Item(item1).Item(item2).Remove "date"
    s.Item(item1).Item(item2).Add arr(i, IBD), "date"
End With

End Function

Function calculs_tableaux_analytiques(portefeuille, i)

With BSH

Dim dansC As Integer
dansC = 0
For Each x In c.Item(portefeuille)
    If x.Item("code") = .Cells(i, IBCC) Then
        dansC = 1
        ''''Tableaux analytiques''''
        '''''''''solde''''''''''''''
        j = c.Item(portefeuille).Item(.Cells(i, IBCC)).Item("solde fin")
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Remove "solde fin"
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add j - .Cells(i, IBV), "solde fin"
        '''''''''autres postes''''''
        If .Cells(i, IBP) = "AACTION" Or .Cells(i, IBP) = "ASIC" Then
            j = c.Item(portefeuille).Item(.Cells(i, IBCC)).Item("acquisition")
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Remove "acquisition"
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Add j - .Cells(i, IBV), "acquisition"
        End If
        If .Cells(i, IBP) = "VACTION" Or .Cells(i, IBP) = "VSIC" Then
            j = c.Item(portefeuille).Item(.Cells(i, IBCC)).Item("cession")
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Remove "cession"
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Add j - .Cells(i, IBV), "cession"
        End If
        If .Cells(i, IBP) = "SCESSA" Or .Cells(i, IBP) = "ECESSA" Then
            j = c.Item(portefeuille).Item(.Cells(i, IBCC)).Item("reclassement")
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Remove "reclassement"
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Add j - .Cells(i, IBV), "reclassement"
        End If
'        If .Cells(i, ibp) = "AUGCAPITAL" Then   'A vérifier
'            j = c.Item(portefeuille).Item(.Cells(i, ibcc)).Item("augmentation capital")
'            c.Item(portefeuille).Item(.Cells(i, ibcc)).Remove "augmentation capital"
'            c.Item(portefeuille).Item(.Cells(i, ibcc)).Add j - .Cells(i, ibv), "augmentation capital"
'        End If
        If .Cells(i, IBP) = "REDCAPITAL" Then
            j = c.Item(portefeuille).Item(.Cells(i, IBCC)).Item("réduction capital")
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Remove "réduction capital"
            c.Item(portefeuille).Item(.Cells(i, IBCC)).Add j - .Cells(i, IBV), "réduction capital"
        End If
        ''''''''''''''''''''''''''''
    End If
Next x

If dansC = 0 Then
    c.Item(portefeuille).Add New Collection, .Cells(i, IBCC)
    c.Item(portefeuille).Item(.Cells(i, IBCC)).Add .Cells(i, IBCC), "code"
    c.Item(portefeuille).Item(.Cells(i, IBCC)).Add .Cells(i, IBCC), "nom"
    c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "solde"
    c.Item(portefeuille).Item(.Cells(i, IBCC)).Add -.Cells(i, IBV), "solde fin"
    'c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "solde fin"
    If .Cells(i, IBP) = "AACTION" Or .Cells(i, IBP) = "ASIC" Then
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add -.Cells(i, IBV), "acquisition"
    Else
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "acquisition"
    End If
    If .Cells(i, IBP) = "VACTION" Or .Cells(i, IBP) = "VSIC" Then
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add -.Cells(i, IBV), "cession"
    Else
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "cession"
    End If
    If .Cells(i, IBP) = "SCESSA" Or .Cells(i, IBP) = "ECESSA" Then
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add -.Cells(i, IBV), "reclassement"
    Else
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "reclassement"
    End If
    If .Cells(i, IBP) = "REDCAPITAL" Then
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add -.Cells(i, IBV), "réduction capital"
    Else
        c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "réduction capital"
    End If
    c.Item(portefeuille).Item(.Cells(i, IBCC)).Add 0, "augmentation capital"
End If

End With

End Function
Function calcul_tableau_analytique_prets(i, valeur As Double, code As String)

Dim dansC As Integer
dansC = 0
For Each x In c
    If x.Item("code") = code Then
        dansC = 1
        ''''Tableaux analytiques''''
        '''''''''solde''''''''''''''
        j = c.Item(code).Item("solde fin")
        c.Item(code).Remove "solde fin"
        c.Item(code).Add j + valeur, "solde fin"
        '''''''''autres postes''''''
        If valeur > 0 Then
            j = c.Item(code).Item("octroi")
            c.Item(code).Remove "octroi"
            c.Item(code).Add j + valeur, "octroi"
        End If
        If valeur < 0 Then
            j = c.Item(code).Item("remboursement")
            c.Item(code).Remove "remboursement"
            c.Item(code).Add j + valeur, "remboursement"
        End If
        ''''''''''''''''''''''''''''
    End If
Next x

If dansC = 0 Then
    c.Add New Collection, code
    c.Item(code).Add code, "code"
    c.Item(code).Add code, "nom"
    c.Item(code).Add 0, "solde"
    c.Item(code).Add valeur, "solde fin"
    If valeur > 0 Then
        c.Item(code).Add valeur, "octroi"
    Else
        c.Item(code).Add 0, "octroi"
    End If
    If valeur < 0 Then
        c.Item(code).Add valeur, "remboursement"
    Else
        c.Item(code).Add 0, "remboursement"
    End If
    c.Item(code).Add 0, "rachat"
    c.Item(code).Add 0, "conversion capital"
End If


End Function

Function correction_fin(item1, item2, sh, ligne, c, cs, cd, cp, cm, ca, cc, rl, rle)
'c: Colonne date            'cs: Colonne solde              'cd: Colonne durée
'cp: Colonne pondération    'cm: Colonne mouvement          'ca: Colonne acquisition
'cc: Colonne cession        'rl: Récap ligne solde          'rle: Récap ligne encours

'Recap: Dernier mois
d = s.Item(item1).Item(item2).Item("date")
RSH.Cells(rl, IRC + Month(d)) = s.Item(item1).Item(item2).Item("solde")

With sh
    ligne2 = IEL + ligne
    'Durèe et solde * durèe (tableau des encours): dernière opération
    du = DateSerial(Year(dd), Month(dd) + 1, 0) - CDate(d) + 1
    .Cells(ligne2, cd) = du
    .Cells(ligne2, cp) = du * s.Item(item1).Item(item2).Item("solde")
    .Cells(ligne2, cp).NumberFormat = "#,##0.00"
    .Cells(ligne2, cm) = s.Item(item1).Item(item2).Item("Mvt")
    .Cells(ligne2, cm).NumberFormat = "#,##0.00"
    If s.Item(item1).Item(item2).Item("Mvt") > 0 Then
        .Cells(ligne2, cc) = s.Item(item1).Item(item2).Item("Mvt")
        .Cells(ligne2, cc).NumberFormat = "#,##0.00"
    Else
        .Cells(ligne2, ca) = -s.Item(item1).Item(item2).Item("Mvt")
        .Cells(ligne2, ca).NumberFormat = "#,##0.00"
    End If

    'Encours Dernière opération
    .Cells(ligne2, c) = s.Item(item1).Item(item2).Item("date")
    .Cells(ligne2, cs) = s.Item(item1).Item(item2).Item("solde")
    .Cells(ligne2, cs).NumberFormat = "#,##0.00"
End With

    'Somme: Solde * Durée
    j = s.Item(item1).Item(item2).Item("Pond.") + du * s.Item(item1).Item(item2).Item("solde")
    s.Item(item1).Item(item2).Remove "Pond."
    s.Item(item1).Item(item2).Add j, "Pond."

    'Encours moyen
    RSH.Cells(rle, IRC + Month(dd)) = _
    s.Item(item1).Item(item2).Item("Pond.") / _
    (DateSerial(Year(dd), Month(dd) + 1, 0) - (CDate("01/01/" & Year(BSH.Cells(2, IBD))) - 1))

End Function

Function correction_fin_prets(sh As Worksheet, dt As Date, valeur As Double, ligne As Integer, y)

'Recap: Dernier mois
d = s.Item("date")
RSH.Cells(IRP, IRC + Month(d)) = s.Item("solde")

With sh
    ligne2 = IEL + ligne
    'Durèe et solde * durèe (tableau des encours): dernière opération
    du = DateSerial(Year(dd), Month(dd) + 1, 0) - CDate(d) + 1
    .Cells(ligne2, 7) = du
    .Cells(ligne2, 8) = du * s.Item("solde")
    .Cells(ligne2, 8).NumberFormat = "#,##0.00"
    .Cells(ligne2, 3) = s.Item("Mvt")
    .Cells(ligne2, 3).NumberFormat = "#,##0.00"
    If s.Item("Mvt") > 0 Then
        .Cells(ligne2, 4) = s.Item("Mvt")
        .Cells(ligne2, 4).NumberFormat = "#,##0.00"
    Else
        .Cells(ligne2, 5) = -s.Item("Mvt")
        .Cells(ligne2, 5).NumberFormat = "#,##0.00"
    End If

    'Encours Dernière opération
    .Cells(ligne2, 2) = s.Item("date")
    .Cells(ligne2, 6) = s.Item("solde")
    .Cells(ligne2, 6).NumberFormat = "#,##0.00"
End With

'Somme: Solde * Durée
j = s.Item("Pond.") + du * s.Item("solde")
s.Remove "Pond."
s.Add j, "Pond."

'Encours moyen
RSH.Cells(IRPE, IRC + Month(dd)) = s.Item("Pond.") / _
(DateSerial(Year(dd), Month(dd) + 1, 0) - (CDate("01/01/" & y) - 1))
'rsh.Cells(rle, irc + Month(dd)) = s.Item("Pond.") / _
(DateSerial(Year(dd), Month(dd) + 1, 0) - (CDate("01/01/" & Year(bsh.Cells(2, ibd))) - 1))

End Function

Sub trier_par_dates(sh, col)

    If Not sh.AutoFilterMode Then
        sh.Range("A1").AutoFilter
    End If

    sh.AutoFilter.Sort.SortFields.Clear
    sh.AutoFilter.Sort.SortFields.Add Key _
        :=Range(col + "2:" + col + "50000"), SortOn:=xlSortOnValues, order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    With sh.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'ActiveWindow.SmallScroll Down:=60

End Sub


Sub réinitialiser()

'variables

'Réinitialiser la feuille 'Récap' (champs titres)
'RSH.Range("E9:P13").ClearContents
'RSH.Range("E18:P22").ClearContents
RSH.Range(IRRE).ClearContents
RSH.Range(IRRS).ClearContents

'Réinitialiser les feuille 'EM Actions', 'EM Trans' et 'EM Tx'
ASH.Cells.ClearContents
YSH.Cells.ClearContents
TSH.Cells.ClearContents
ASH.Cells.ClearFormats
YSH.Cells.ClearFormats
TSH.Cells.ClearFormats

'Réinitialiser la feuille 'Tableaux analytiques'
TASH.Cells.ClearContents
TASH.Cells.ClearFormats
TASH.Cells.Interior.Color = xlNone

End Sub
Sub réinitialiser_prets()

'variables

'Réinitialiser la feuille 'Récap' (champs prêts)
'RSH.Range("E14:P14").ClearContents
'RSH.Range("E23:P23").ClearContents
RSH.Range(IRREP).ClearContents
RSH.Range(IRRSP).ClearContents

'Réinitialiser les feuille 'EM Prêts'
PSH.Cells.ClearContents
PSH.Cells.ClearFormats
PSH.Cells.Interior.Color = xlNone

'Réinitialiser la feuille 'Tableau Analytique Prêts'
TPSH.Cells.ClearContents
TPSH.Cells.ClearFormats
TPSH.Cells.Interior.Color = xlNone

End Sub
Sub optimisationDébut()

Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
'Application.Calculation = xlManual
Application.ScreenUpdating = False

End Sub


Sub optimisationFin()

Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Sub init_tableaux_analytiques()

'Dim titres As String * 100
Dim lignetitre As Integer
Dim arr() As Variant
Dim titres() As Variant

Const itat As Integer = 2   'Colonne titre
Const itad As Integer = 3   'Colonne solde début
Const itaa As Integer = 4   'Colonne acquisition

Dim itaak As Integer       'Colonne augmentation capital
Dim itar As Integer        'Colonne reclassement
Dim itark As Integer       'Colonne réduction capital
Dim itac As Integer        'Colonne cession
Dim itasf As Integer       'Colonne solde fin

With TASH

'Dessine les tableaux analytiques (noms + codes des titres)
x = IEL - 2
For Each i In Split("TRANS PART PLACT")

    'titres = "Titre,Début,Acquisition,Reclassement,Cession,Fin"
    titres = Array("Titre", "Solde Début", "Acquisition", "Reclassement", "Cession", "Solde Fin")
    'index tableaux analytiques
    'itaa = 4        'Colonne acquisition
    itar = 5        'Colonne reclassement
    itac = 6        'Colonne cession
    itasf = 7       'Colonne solde fin
    
    If i = "PART" Then
        'titres = "Titre,Début,Acquisition,Aug. Capital,Reclassement,Red. Capital,Cession,Fin"
        titres = Array("Titre", "Solde Début", "Acquisition", "Aug. Capital", "Reclassement", "Red. Capital", "Cession", "Solde Fin")
        itaak = 5       'Colonne augmentation capital
        itar = 6        'Colonne reclassement
        itark = 7       'Colonne réduction capital
        itac = 8        'Colonne cession
        itasf = 9       'Colonne solde fin
    End If
       
    'Grand titre
    .Cells(x, 2) = i        'Nom du portefeuille (titre)
'    With TASH.Cells(x, 2)
'        .Font.Bold = True
'        .HorizontalAlignment = xlCenter
'        .Interior.Color = RGB(64, 64, 64)
'        .Font.Color = vbWhite
'    End With
'    .Range(.Cells(x, 2), .Cells(x, UBound(titres) + 2)).Merge
    
    'Sous-titres (array)
    .Range(.Cells(x + 1, 2), .Cells(x + 1, UBound(titres) + 2)) = titres
'    With TASH.Range(.Cells(x + 1, 2), .Cells(x + 1, UBound(titres) + 2))
'        .Font.Bold = True
'        .HorizontalAlignment = xlCenter
'        .Interior.Color = RGB(83, 142, 213)
'        .Font.Color = vbWhite
'    End With
        
    li = 1
    x = x + 2
    total_debut = x
    'Lignes du tableau
    For Each j In c.Item(i)
        .Cells(x, itat) = j.Item("nom")
        .Cells(x, itad) = j.Item("solde")
        .Cells(x, itasf) = j.Item("solde fin")
        .Cells(x, itaa) = j.Item("acquisition")
        .Cells(x, itar) = j.Item("reclassement")
        .Cells(x, itac) = j.Item("cession")
        If i = "PART" Then
            .Cells(x, itaak) = j.Item("augmentation capital")
            .Cells(x, itark) = j.Item("réduction capital")
            .Cells(x, itaak).NumberFormat = "#,##0.00"
            .Cells(x, itark).NumberFormat = "#,##0.00"
        End If
        
        .Cells(x, 3).NumberFormat = "#,##0.00"
        .Cells(x, itasf).NumberFormat = "#,##0.00"
        .Cells(x, itaa).NumberFormat = "#,##0.00"
        .Cells(x, itar).NumberFormat = "#,##0.00"
        .Cells(x, itac).NumberFormat = "#,##0.00"
        'If li Mod 2 = 1 Then colorer_ligne_tableau TASH, Int(x), 2, Int(itasf)
        li = li + 1
        x = x + 1
    Next j
    
    
    colorer_tableau TASH, Int(total_debut), itat, itasf
    
    'Ajout des totaux
    .Cells(x, 2) = "Total"
    .Cells(x, 2).Font.Bold = True
    For j = 3 To itasf
        .Cells(x, j).Formula = "=sum(" + .Cells(total_debut, j).Address + ":" + .Cells(x - 1, j).Address + ")"
        .Cells(x, j).Font.Bold = True
    Next j
    .Range(.Cells(x, 2), .Cells(x, itasf)).Interior.Color = RGB(64, 64, 64)
    .Range(.Cells(x, 2), .Cells(x, itasf)).Font.Color = vbWhite
    
    x = x + 2
Next i
End With

End Sub
Sub init_tableau_analytique_prets()

Const itate As Integer = 2    'Colonne titre emprunteur
Const itasd As Integer = 3    'Colonne solde début
Const itacc As Integer = 4    'Colonne conversion en capital
Const itanp As Integer = 5    'Colonne nouveaux prêts
Const itare As Integer = 6    'Colonne remboursements
Const itara As Integer = 7    'Colonne rachats
Const itasf As Integer = 8    'Colonne solde fin

With TPSH
    x = IEL - 2
    titres = Array("Emprunteur", "Solde Début", "Conversion en Capital", "Nouveaux Prêts", "Remboursements", "Rachat", "Solde Fin")
    .Cells(x, itate) = "PRETS"        'Nom du portefeuille (titre)
    .Range(.Cells(x + 1, itate), .Cells(x + 1, UBound(titres) + itate)) = titres
    
    x = x + 2
    total_debut = x
    'Lignes du tableau
    For Each j In c
        .Cells(x, itate) = j.Item("nom")
        .Cells(x, itasd) = j.Item("solde")
        .Cells(x, itasf) = j.Item("solde fin")
        .Cells(x, itanp) = j.Item("octroi")
        .Cells(x, itare) = j.Item("remboursement")
        .Cells(x, itara) = j.Item("rachat")
        .Cells(x, itacc) = j.Item("conversion capital")
        .Cells(x, itasd).NumberFormat = "#,##0.00"
        .Cells(x, itasf).NumberFormat = "#,##0.00"
        .Cells(x, itanp).NumberFormat = "#,##0.00"
        .Cells(x, itare).NumberFormat = "#,##0.00"
        .Cells(x, itara).NumberFormat = "#,##0.00"
        .Cells(x, itacc).NumberFormat = "#,##0.00"
        x = x + 1
    Next j
    
    colorer_tableau TPSH, Int(total_debut), itate, itasf
    
    'Ajout des totaux
    .Cells(x, 2) = "Total"
    .Cells(x, 2).Font.Bold = True
    For j = 3 To itasf
        .Cells(x, j).Formula = "=sum(" + .Cells(total_debut, j).Address + ":" + .Cells(x - 1, j).Address + ")"
        .Cells(x, j).Font.Bold = True
    Next j
    .Range(.Cells(x, 2), .Cells(x, itasf)).Interior.Color = RGB(64, 64, 64)
    .Range(.Cells(x, 2), .Cells(x, itasf)).Font.Color = vbWhite

End With


End Sub

Sub encours_prets_2()

'''' Enlever le commentaire pour les tests ''''
'variables
'''''''''''''''''''''''''''''''''''''''''''''''

trier_par_dates BPSH, "E"

'Dates début et fin
dd = CDate(RSH.Range(IRPDD).Value)
df = CDate(RSH.Range(IRPDF).Value)

'Détermine l'année de base (peut générer des bugs)
Dim y As Integer
'y = Year(bpsh.Cells(2, ipbr)) 'Peut générer un bug s'il n'y a que des prêts linéaires avec des dates différentes de l'anée actuelle)
y = Year(dd)

'Collection d'opérations
Set o = New Collection

'Collection tableau analytique
Set c = New Collection

'Crée le contenant des calculs
Set s = New Collection
s.Add CDate("01/01/" & y), "date"       'Date
s.Add 0, "Mvt"          'Mouvement
s.Add 0, "Pond."        'Solde * Durée

'Solde de départ
With SSH
i = ISC
Do Until IsEmpty(.Cells(ISL, i))
    If y - 1 = .Cells(ISL, i) Then s.Add .Cells(ISLP, i), "solde"
    i = i + 1
Loop
End With

'Cherche les soldes de départ du tableau analytique
With SSH
'Cherche la colonne correspondante à l'année de base dans le tableau des soldes
i = isac
Do Until IsEmpty(.Cells(ISAL, i))
    If y - 1 = .Cells(ISAL, i) Then isac = i
    i = i + 1
Loop
'Compile la liste des prêts
i = ISAL + 1
Do Until IsEmpty(.Cells(i, ISACN))
    If .Cells(i, ISACP) = "PRETS" Then
        c.Add New Collection, .Cells(i, ISACC)
        c.Item(.Cells(i, ISACC)).Add .Cells(i, ISACC), "code"
        c.Item(.Cells(i, ISACC)).Add .Cells(i, ISACN), "nom"
        c.Item(.Cells(i, ISACC)).Add .Cells(i, isac), "solde"
        c.Item(.Cells(i, ISACC)).Add .Cells(i, isac), "solde fin"
        c.Item(.Cells(i, ISACC)).Add 0, "octroi"
        c.Item(.Cells(i, ISACC)).Add 0, "remboursement"
        c.Item(.Cells(i, ISACC)).Add 0, "rachat"
        c.Item(.Cells(i, ISACC)).Add 0, "conversion capital"
    End If
    i = i + 1
Loop
End With

'Extraction des numéros d'opération des remboursements linéaires
With ESH
Dim arr As Variant
i = 2
codes = .Cells(i, IEN)
Do Until IsEmpty(.Cells(i, IEN))
    arr = Split(codes)
    If CDbl(.Cells(i, IEN)) <> CDbl(arr(UBound(arr))) Then codes = codes & " " & .Cells(i, IEN)
    i = i + 1
Loop
End With

'Boucle sur la base de données
With BPSH
i = 2
x = 0
Do Until IsEmpty(.Cells(i, 3))
    'Ignorer EMLT
    If .Cells(i, IPBP) <> "EMLT" Then
        If CDate(.Cells(i, IPBD)) >= dd And CDate(.Cells(i, IPBD)) <= df Then
            'Augmentation stock de prêts
            xx = str(x)
            o.Add New Collection, xx
            o.Item(xx).Add .Cells(i, IPBC), "code"
            o.Item(xx).Add CDate(.Cells(i, IPBD)), "date"
            o.Item(xx).Add CDbl(.Cells(i, IPBN)), "valeur"
            x = x + 1
        End If
        lin = 0
        For Each j In Split(codes)
            If CDbl(.Cells(i, IPBO)) = CDbl(j) Then lin = 1
        Next j
        If lin = 1 Then
        'Remboursement linéaire (traitement à part)
            k = 2
            Do Until IsEmpty(ESH.Cells(k, IEN))
                'If .Cells(i, ipbo) = esh.Cells(k, ien) And dd <= CDate(esh.Cells(k, ied)) <= df Then
                If .Cells(i, IPBO) = ESH.Cells(k, IEN) _
                And dd <= CDate(ESH.Cells(k, IED)) _
                And CDate(ESH.Cells(k, IED)) <= df Then
                    xx = str(x)
                    o.Add New Collection, xx
                    o.Item(xx).Add .Cells(i, IPBC), "code"
                    o.Item(xx).Add CDate(ESH.Cells(k, IED)), "date"
                    o.Item(xx).Add -CDbl(ESH.Cells(k, IET)), "valeur"
                    x = x + 1
                End If
                k = k + 1
            Loop
        End If
        If lin = 0 And CDate(.Cells(i, IPBR)) >= dd And CDate(.Cells(i, IPBR)) <= df Then
        'Remboursement infine (Traitement normal)
            xx = str(x)
            o.Add New Collection, xx
            o.Item(xx).Add .Cells(i, IPBC), "code"
            o.Item(xx).Add CDate(.Cells(i, IPBR)), "date"
            o.Item(xx).Add -CDbl(.Cells(i, IPBV)), "valeur"
            x = x + 1
        End If
    End If
    i = i + 1
Loop
End With

'Tri des opérations par dates
'Insertion sort
'x = 0
For x = 1 To o.Count - 1
    j = x
    Do While j >= 1
        If o.Item(str(j)).Item("date") <= o.Item(str(j - 1)).Item("date") Then
            Set a = o.Item(str(j - 1))
            o.Remove str(j - 1)
            o.Add o.Item(str(j)), str(j - 1)
            o.Remove str(j)
            o.Add a, str(j)
        End If
        j = j - 1
    Loop
Next x

'''' Début du calcul des encours moyens ''''''''''''
'affiche_titres_encours_op_prets psh, iec, "PRETS"
affiche_titres_encours_op PSH, IEC, "PRETS"
dd = o.Item(str(0)).Item("date")      'Date de la dernière opération
k = 0
For i = 0 To o.Count - 1

    If Month(dd) = Month(o.Item(str(i)).Item("date")) - 1 Then    'Solde mensuel
        '''' <- Calcul mensuel
        calcul_mensuel_prets IRP, IRPE, y
    End If
    
    'Détermine la date de la dernière opération
    If CDate(o.Item(str(i)).Item("date")) > CDate(dd) Then
        dd = o.Item(str(i)).Item("date")
    End If
    
    '''' <- Calcul par opération
    d = s.Item("date")
    If CDate(d) <> CDate(dd) Then
        k = k + 1
    End If

    calcul_prets_3 PSH, CDate(o.Item(str(i)).Item("date")), CDbl(o.Item(str(i)).Item("valeur")), Int(k - 1)
    'If k Mod 2 = 1 Then colorer_ligne_tableau PSH, IEL + k - 1, IEC, IECP
    calcul_tableau_analytique_prets i, CDbl(o.Item(str(i)).Item("valeur")), o.Item(str(i)).Item("code")
Next i

i = o.Count - 1
correction_fin_prets PSH, CDate(o.Item(str(i)).Item("date")), CDbl(o.Item(str(i)).Item("valeur")), Int(k), y

'For i = 0 To o.Count - 1
'    xsh.Cells(i + 2, 1) = o.Item(str(i)).Item("date")
'    xsh.Cells(i + 2, 2) = o.Item(str(i)).Item("code")
'    xsh.Cells(i + 2, 3) = o.Item(str(i)).Item("valeur")
'Next i

'Calcul par titre (tableau analytique)

colorer_tableau PSH, IEL, IEC, IECP


End Sub
Function colorer_ligne_tableau(sh, ligne As Integer, min As Integer, max As Integer)
    
'Colorer les lignes du tableau
With sh
    'r = .Cells(iel + k - 1, min).Address & ":" & .Cells(iel + k - 1, max).Address
    'r = .Cells(ligne, min).Address & ":" & .Cells(ligne, max).Address
    '.Range(r).Interior.Color = RGB(216, 216, 216)
    
    .Range(.Cells(ligne, min), .Cells(ligne, max)).Interior.Color = RGB(216, 216, 216)
End With



End Function

Function colorer_tableau(sh As Worksheet, ligne As Integer, min As Integer, max As Integer)

Dim ligneGrandTitre As Integer
Dim ligneSousTitre As Integer

ligneGrandTitre = ligne - 2
ligneSousTitre = ligne - 1

With sh
    'Colore le grand titre
    With sh.Cells(ligneGrandTitre, min)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = vbWhite
    End With
    .Range(.Cells(ligneGrandTitre, min), .Cells(ligneGrandTitre, max)).Merge

    'Colore les sous-titres
    With sh.Range(.Cells(ligneSousTitre, min), .Cells(ligneSousTitre, max))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(83, 142, 213)
        .Font.Color = vbWhite
    End With
    'Colore les lignes
    k = 0
    Do Until IsEmpty(.Cells(ligne + k, min))
        If k Mod 2 = 0 Then .Range(.Cells(ligne + k, min), .Cells(ligne + k, max)).Interior.Color = RGB(216, 216, 216) 'Gris clair
        k = k + 1
    Loop
End With


End Function
Function calcul_prets_3(sh As Worksheet, dt As Date, valeur As Double, ligne As Integer)

With sh
    ligne2 = IEL + ligne
    d = s.Item("date")
    If d <> dt Then
        .Cells(ligne2, 2) = s.Item("date")
        .Cells(ligne2, 6) = s.Item("solde")
        .Cells(ligne2, 6).NumberFormat = "#,##0.00"
        du = CDate(dt) - CDate(d)
        .Cells(ligne2, 7) = du
        .Cells(ligne2, 8) = du * s.Item("solde")
        .Cells(ligne2, 8).NumberFormat = "#,##0.00"
        .Cells(ligne2, 3) = s.Item("Mvt")
        .Cells(ligne2, 3).NumberFormat = "#,##0.00"
        If s.Item("Mvt") > 0 Then
            .Cells(ligne2, 4) = s.Item("Mvt")
            .Cells(ligne2, 4).NumberFormat = "#,##0.00"
        Else
            .Cells(ligne2, 5) = -s.Item("Mvt")
            .Cells(ligne2, 5).NumberFormat = "#,##0.00"
        End If
        s.Remove "Mvt"
        s.Add valeur, "Mvt"
        'Somme: Solde * Durée
        j = s.Item("Pond.") + du * s.Item("solde")
        s.Remove "Pond."
        s.Add j, "Pond."
    Else
        j = s.Item("Mvt") + valeur
        s.Remove "Mvt"
        s.Add j, "Mvt"
    End If
    j = s.Item("solde") + valeur
    s.Remove "solde"
    s.Add j, "solde"
    s.Remove "date"
    s.Add dt, "date"
End With

End Function


Function calcul_mensuel_prets(rl, rle, y)
'Function calcul_mensuel_prets(dt As Date, valeur As Double, ligne As Integer)
    RSH.Cells(rl, IRC + Month(dd)) = s.Item("solde")
    'Encours mensuel
    d = s.Item("date")
    dn = DateSerial(Year(dd), Month(dd) + 1, 0) - CDate(d) + 1      'dernière opération du mois -> fin de mois + 1
    j = s.Item("Pond.") + s.Item("solde") * dn
    RSH.Cells(rle, IRC + Month(dd)) = j / _
    (DateSerial(Year(dd), Month(dd) + 1, 0) - _
    (CDate("01/01/" & y) - 1))

End Function
