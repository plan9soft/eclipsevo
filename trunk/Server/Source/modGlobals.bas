Attribute VB_Name = "modGlobals"
Option Explicit

' Server variables
Public SERV_ISRUNNING As Boolean

' Global Variables.
Public GAME_NAME As String
Public WEB_SITE As String
Public MAX_PLAYERS As Integer
Public MAX_CLASSES As Integer
Public MAX_SPELLS As Integer
Public MAX_SCRIPTSPELLS As Integer
Public MAX_ELEMENTS As Integer
Public MAX_MAPS As Integer
Public MAX_SHOPS As Integer
Public MAX_ITEMS As Integer
Public MAX_NPCS As Integer
Public MAX_MAP_ITEMS As Integer
Public MAX_GUILDS As Integer
Public MAX_GUILD_MEMBERS As Integer
Public MAX_PARTY_MEMBERS As Integer
Public MAX_EMOTICONS As Integer
Public MAX_LEVEL As Integer
Public MAX_SERVLINES As Long
Public SCRIPTING As Byte
Public SCRIPT_DEBUG As Byte
Public PAPERDOLL As Byte
Public SPRITESIZE As Byte
Public CUSTOM_SPRITE As Integer
Public PKMINLVL As Integer
Public LEVEL As Integer
Public EMAIL_AUTH As Integer
Public HP_REGEN As Byte
Public HP_TIMER As Long
Public MP_REGEN As Byte
Public MP_TIMER As Long
Public SP_REGEN As Byte
Public SP_TIMER As Long
Public NPC_REGEN As Byte
Public SP_ENABLE As Byte
Public STAT1 As String
Public STAT2 As String
Public STAT3 As String
Public STAT4 As String
Public SAVETIME As Long
Public CLASSES As Byte
Public SP_ATTACK As Byte
Public SP_RUNNING As Byte

' Global Timers.
Public CHATLOG_TIMER As Long
Public SHUTDOWN_TIMER As Long
Public PLYRSAVE_TIMER As Long

' Map Coords.
Public MAX_MAPX As Long
Public MAX_MAPY As Long

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' File paths.
Public FILE_DATAINI As String
Public FILE_STATSINI As String
Public FILE_NEWSINI As String
Public FILE_MOTDINI As String
Public FILE_TILESINI As String

' Folder paths.
Public FLDR_MAPS As String
Public FLDR_LOGS As String
Public FLDR_ACCOUNTS As String
Public FLDR_NPCS As String
Public FLDR_ITEMS As String
Public FLDR_SPELLS As String
Public FLDR_SHOPS As String
Public FLDR_BANKS As String
Public FLDR_CLASSES As String

Public Map() As MapRec
Public MapCache() As String
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public ClassData() As ClassRec
Public Item() As ItemRec
Public NPC() As NpcRec
Public MapItem() As MapItemRec
Public MapNPC() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Emoticons() As EmoRec
Public Element() As ElementRec
Public Party() As NewPartyRec
Public Experience() As Long
Public CTimers As Collection

Public Arrows(1 To MAX_ARROWS) As ArrowRec

Public AddHP As StatRec
Public AddMP As StatRec
Public AddSP As StatRec

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Global PlayerI As Long

' Winsock globals
Public GAME_PORT As Long
Public MAX_PACKETS As Long
Public MAX_BYTES As Long

' Map Control
Public IS_SCROLLING As Long

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public WeatherType As Long
Public GameTime As Long
Public WeatherLevel As Long
Public GameClock As String
Public Gamespeed As Long

Public Hours As Integer
Public Seconds As Long
Public Minutes As Integer

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveMPTimer As Long
Public GiveSPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

Public TimeDisable As Boolean

' VBScript - The VBScript object.
Public MyScript As clsSadScript

' VBScript - The command file.
Public clsScriptCommands As clsCommands

' Our GameServer and Sockets objects
Public GameServer As clsServer
