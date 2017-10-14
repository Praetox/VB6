Attribute VB_Name = "Constants"
Public Light_S As Long
    
    '-- CHAR --
    Public Const CH_ID = &H6059D0
    Public Const CH_HP = &H6059CC
    Public Const CH_Ma = &H6059B0
    Public Const CH_Cap = &H6059A0
    Public Const CH_Lvl = &H6059C0
    Public Const CH_Exp = &H6059C4
    Public Const CH_Sol = &H6059A8
    Public Const CH_Mlv = &H6059BC
    Public Const CH_Clk = &H75D440
    Public Const CH_TSt = &H75EA40
    Public Const CH_TTi = &H75EA3C
    Public Const CH_Con = &H75D3E0
    Public Const CH_X = &H610C28
    Public Const CH_Y = &H610C24
    Public Const CH_Z = &H610C20
    Public Const CH_gX = &H605A14
    Public Const CH_gY = &H605A10
    Public Const CH_gZ = &H605A0C
    Public Const CH_Chm = &H75A9B8
    
    Public Const CH_S0 = &H60E304 'Ammo
    Public Const CH_S1 = &H60E298 'Head
    Public Const CH_S2 = &H60E2BC 'Armor
    Public Const CH_S3 = &H60E2C8 'Right
    Public Const CH_S4 = &H60E2D4 'Left
    Public Const CH_S5 = &H60E2E0 'Legs
    Public Const CH_S6 = &H60E2EC 'Feet
    Public Const CH_S7 = &H60E2A4 'Amulet
    Public Const CH_S8 = &H60E2F8 'Ring
    
    '-- BATTLE LIST --
    Public Const BL_Start = &H605A34
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
    Public Const CT_Start = &H60E310
    Public Const CTD_Container = 492
    Public Const CTD_ContainerID = 4
    Public Const CTD_ContainerItem = 12
    Public Const CTD_ContainerName = 16
    Public Const CTD_ContainerVolume = 48 'Maximum of items in container
    Public Const CTD_ContainerAmount = 56 'Current items in container
    Public Const CTD_ContainerItemID = 60
    Public Const CTD_ContainerItemCount = 64 'Stacked item count
    
    '-- MAP --
    Public Const MAP_POINTER = &H615738
    Public Const Map_TileDist = 172
    Public Const Map_TopObject = 28 '16
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
    Public Const WM_KD = 256
    Public Const WM_KU = 257
    Public Const BOX_1 = &H605994
    Public Const BOX_2 = &H605998
    Public Const BOX_3 = &H60599C
    Public Const Look_ID = &H75D488
    Public Const Look_Ct = &H75D480
    Public Const Look_TX = &H75EC68
    Public Const Ru_Explo = &HC80
    Public Const Ru_GFB = &HC77
    Public Const Ru_HMM = &HC7E
    Public Const Ru_IH = &HC50
    Public Const Ru_SD = &HC53
    Public Const Ru_UH = &HC58
    Public Const Ru_NE = &HC4B

    '-- ITEMS --
    Public Const CONT_BAG = &HB25
    Public Const ITEM_GOLD = &HBD7
    Public Const ITEM_ROPE = &HBBB
    Public Const FOOD_FISH = &HDFA
