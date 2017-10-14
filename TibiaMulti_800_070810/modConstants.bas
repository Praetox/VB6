Attribute VB_Name = "Constants"
Public Light_S As Long
    
    '-- CHAR --
    Public Const CH_ID = &H60EAD0
    Public Const CH_HP = &H60EACC
    Public Const CH_Ma = &H60EAB0
    Public Const CH_Cap = &H60EAA0
    Public Const CH_Lvl = &H60EAC0
    Public Const CH_Exp = &H60EAC4
    Public Const CH_Sol = &H60EAA8
    Public Const CH_Mlv = &H60EABC
    Public Const CH_Clk = &H766E58
    Public Const CH_TSt = &H768458
    Public Const CH_TTi = &H768454
    Public Const CH_Con = &H766DF8
    Public Const CH_X = &H6198F8
    Public Const CH_Y = &H6198F4
    Public Const CH_Z = &H6198F0
    Public Const CH_gX = &H60EB14
    Public Const CH_gY = &H60EB10
    Public Const CH_gZ = &H60EB0C
    
    Public Const CH_S0 = &H616FF4 'Ammo
    Public Const CH_S1 = &H616F88 'Head
    Public Const CH_S2 = &H616FAC 'Armor
    Public Const CH_S3 = &H616FB8 'Right
    Public Const CH_S4 = &H616FC4 'Left
    Public Const CH_S5 = &H616FD0 'Legs
    Public Const CH_S6 = &H616FDC 'Feet
    Public Const CH_S7 = &H616F94 'Amulet
    Public Const CH_S8 = &H616FE8 'Ring
    
    '-- BATTLE LIST --
    Public Const BL_Start = &H60EB34
    Public Const BL_End = BL_Start + (147 * &HA0)
    Public Const BL_Dist = &HA0
        Public Const BL_ID = -4
        Public Const BL_Name = 0
        Public Const BL_X = 32
        Public Const BL_Y = 36
        Public Const BL_Z = 40
        Public Const BL_Wlk = 72
        Public Const BL_Dir = 76
        Public Const BL_OFit = 92
        Public Const BL_HP = 132
        Public Const BL_Spd = 136
        Public Const BL_Vis = 140
        Public Const BL_LStr = 116
        Public Const BL_LCol = 120
    
    '-- OUTFITS --
    Public Const OFit_Invis = 0
    Public Const OFit_GM = 75
    Public Const OFit_Druid = 128
    Public Const OFit_Pally = 129
    Public Const OFit_Sorc = 130
    Public Const OFit_Knight = 131
    
    '-- CONTAINERS --
    Public Const CT_Start = &H617000
    Public Const CTD_Container = 492
    Public Const CTD_ContainerID = 4
    Public Const CTD_ContainerItem = 12
    Public Const CTD_ContainerName = 16
    Public Const CTD_ContainerVolume = 48 'Max. items in container
    Public Const CTD_ContainerAmount = 56 'Amount of items in container
    Public Const CTD_ContainerItemID = 60
    Public Const CTD_ContainerItemCount = 64 'Amount of stacked items
    
    '-- MAP --
    Public Const MAP_POINTER = &H61E408 '&H61E838
    Public Const Map_TileDist = 172
    Public Const Map_ObjectDist = 12
    Public Const Map_ObjectIdDist = 0
    Public Const Map_ObjectDataDist = 4
    Public Const Map_ObjectIdOffset = 4
    Public Const Map_ObjectDataOffset = 8
    
    '-- TILES --
    Public Const TILE_LADDER_HOLE = 411
    Public Const TILE_TRANSPARENT = 470
    Public Const TILE_WATER_OLD = 865
    Public Const TILE_WATER_FISH_BEGIN = 4597
    Public Const TILE_WATER_FISH_END = 4602
    Public Const TILE_WATER_NOFISH_BEGIN = 4603
    Public Const TILE_WATER_NOFISH_END = 4614
    
    '-- OTHERS --
    Public Const ID_Look = &H766EA0
    Public Const ID_Clik = &H766E94
    Public Const WM_KD = 256
    Public Const WM_KU = 257
    Public Const BOX_1 = &H60EA94
    Public Const BOX_2 = &H60EA98
    Public Const BOX_3 = &H60EA9C
    Public Const Ru_Explo = &HC80
    Public Const Ru_GFB = &HC77
    Public Const Ru_HMM = &HC7E
    Public Const Ru_IH = &HC50
    Public Const Ru_SD = &HC53
    Public Const Ru_UH = &HC58
    Public Const Ru_NE = &HC4B

    '-- ITEMS --
    Public Const ITEN_GOLD = &HBD7
    Public Const ITEN_ROPE = &HBBB
    Public Const FOOD_FISH = &HDFA
