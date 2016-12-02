import xlrd
from xlutils.copy import copy

import Abaddon as h1
import Alchemist as h2
import AncientApparition as h3
import AntiMage as h4
import ArcWarden as h5
import Axe as h6
import Bane as h7
import Batrider as h8
import Beastmaster as h9
import Bloodseeker as h10
import BountyHunter as h11
import BrewMaster as h12
import BristleBack as h13
import BroodMother as h14
import CentaurWarrunner as h15
import ChaosKnight as h16
import Chen as h17
import Clink as h18
import ClockWerk as h19
import CrystalMaiden as h20
import DarkSeer as h21
import Dazzle as h22
import DeathProphet as h23
import Disruptor as h24
import Doom as h25
import DragonKnight as h26
import DrowRanger as h27
import EarthShaker as h28
import EarthSpirit as h29
import ElderTitan as h30
import EmberSpirit as h31
import Enchantress as h32
import Enigma as h33
import FacelessVoid as h34
import Gyrocopter as h35
import Huskar as h36
import Invoker as h37
import Io as h38
import Jakiro as h39
import Juggernaut as h40
import KeeperOftheLight as h41
import Kunkka as h42
import LegionCommander as h43
import Leshrac as h44
import Lich as h45
import Lifestealer as h46
import Lina as h47
import Lion as h48
import LoneDruid as h49
import Luna as h50
import Lycan as h51
import Magnus as h52
import Medusa as h53
import Meepo as h54
import Mirana as h55
import Morpling as h56
import NagaSiren as h57
import NaturesProphet as h58
import Necrophos as h59
import NightStalker as h60
import NyxAssasin as h61
import OgreMagi as h62
import Omniknight as h63
import Oracle as h64
import OutworldDevourer as h65
import PhantomAssasin as h66
import PhantomLancer as h67
import Phoenix as h68
import Puck as h69
import Pudge as h70
import Pugna as h71
import QueenofPain as h72
import Razor as h73
import Riki as h74
import Rubick as h75
import SandKing as h76
import ShadowDemon as h77
import ShadowFiend as h78
import ShadowShaman as h79
import Silencer as h80
import SkywrathMage as h81
import Slardar as h82
import Slark as h83
import Sniper as h84
import Spectre as h85
import SpiritBreaker as h86
import StormSpirit as h87
import Sven as h88
import Techies as h89
import TemplarAssasin as h90
import Terrorblade as h91
import Tidehunter as h92
import Timbersaw as h93
import Tinker as h94
import Tiny as h95
import TreantProtector as h96
import TrollWarlord as h97
import Tusk as h98
import Underlord as h99
import Undying as h100
import Ursa as h101
import VengefulSpirit as h102
import Venomancer as h103
import Viper as h104
import Visage as h105
import Warlock as h106
import Weaver as h107
import Windranger as h108
import WinterWyvern as h109
import WitchDoctor as h110
import WraithKing as h111
import Zeus as h112
def write_data(key , data):
    """write dara to xls"""
##    key = "Axe"
##    data = da.statistic
    wb = xlrd.open_workbook('Data.xls')
    s = wb.sheet_by_index(0)
    ww = copy(wb)
    s0 = ww.get_sheet(0)
    s1 = ww.get_sheet(1)
    s2 = ww.get_sheet(2)
    for y in range(1,113):
        if s.cell_value(0,y) == key:
            key = y
            break
    i = 1
    for x in range(1,113):
        if s.cell_value(0,key) != s.cell_value(x,0):
            s0.write(y,x,float(str(data[x-i][1])[:-1]))
            s1.write(y,x,float(str(data[x-i][2])[:-1]))
            s2.write(y,x,(str(data[x-i][3])))
        else:
            i = 2
    ww.save('Data.xls')
##write_data()
write_data("Abaddon",h1.statistic)
write_data("Alchemist",h2.statistic)
write_data("Ancient Apparition",h3.statistic)
write_data("Anti-Mage",h4.statistic)
write_data("Arc Warden",h5.statistic)
write_data("Axe",h6.statistic)
write_data("Bane",h7.statistic)
write_data("Batrider",h8.statistic)
write_data("Beastmaster",h9.statistic)
write_data("Bloodseeker",h10.statistic)
write_data("Bounty Hunter",h11.statistic)
write_data("Brewmaster",h12.statistic)
write_data("Bristleback",h13.statistic)
write_data("Broodmother",h14.statistic)
write_data("Centaur Warrunner",h15.statistic)
write_data("Chaos Knight",h16.statistic)
write_data("Chen",h17.statistic)
write_data("Clinkz",h18.statistic)
write_data("Clockwerk",h19.statistic)
write_data("Crystal Maiden",h20.statistic)
write_data("Darsk Seer",h21.statistic)
write_data("Dazzle",h22.statistic)
write_data("Death Prophet",h23.statistic)
write_data("Distuptor",h24.statistic)
write_data("Doom",h25.statistic)
write_data("Dragon Knight",h26.statistic)
write_data("Drow Ranger",h27.statistic)
write_data("Earth Spirit",h28.statistic)
write_data("Earthshaker",h29.statistic)
write_data("Elder Titan",h30.statistic)
write_data("Ember Spirit",h31.statistic)
write_data("Enchantress",h32.statistic)
write_data("Enigma",h33.statistic)
write_data("Faceless Void",h34.statistic)
write_data("Gyrocopter",h35.statistic)
write_data("Huskar",h36.statistic)
write_data("Invoker",h37.statistic)
write_data("Io",h38.statistic)
write_data("Jakiro",h39.statistic)
write_data("Juggernaut",h40.statistic)
write_data("Keeper of the Light",h41.statistic)
write_data("Kunkka",h42.statistic)
write_data("Legion Commander",h43.statistic)
write_data("Leshrac",h44.statistic)
write_data("Lich",h45.statistic)
write_data("Lifestealer",h46.statistic)
write_data("Lina",h47.statistic)
write_data("Lion",h48.statistic)
write_data("Lone Druid",h49.statistic)
write_data("Luna",h50.statistic)
write_data("Lycan",h51.statistic)
write_data("Magnus",h52.statistic)
write_data("Medusa",h53.statistic)
write_data("Meepo",h54.statistic)
write_data("Mirana",h55.statistic)
write_data("Morphling",h56.statistic)
write_data("Naga Siren",h57.statistic)
write_data("Nature's Prophet",h58.statistic)
write_data("Necrophos",h59.statistic)
write_data("Night Stalker",h60.statistic)
write_data("Nyx Assasin",h61.statistic)
write_data("Ogre Magi",h62.statistic)
write_data("Omniknight",h63.statistic)
write_data("Oracle",h64.statistic)
write_data("Outworld Devourer",h65.statistic)
write_data("Phantom Assasin",h66.statistic)
write_data("Phantom Lancer",h67.statistic)
write_data("Phoenix",h68.statistic)
write_data("Puck",h69.statistic)
write_data("Pudge",h70.statistic)
write_data("Pugna",h71.statistic)
write_data("Queen of Pain",h72.statistic)
write_data("Razor",h73.statistic)
write_data("Riki",h74.statistic)
write_data("Rubick",h75.statistic)
write_data("Sand King",h76.statistic)
write_data("Shadow Demon",h77.statistic)
write_data("Shadow Fiend",h78.statistic)
write_data("Shadow Shaman",h79.statistic)
write_data("Silencer",h80.statistic)
write_data("Skywrath Mage",h81.statistic)
write_data("Slardar",h82.statistic)
write_data("Slark",h83.statistic)
write_data("Sniper",h84.statistic)
write_data("Spectre",h85.statistic)
write_data("Spirit Breaker",h86.statistic)
write_data("Storm Spirit",h87.statistic)
write_data("Sven",h88.statistic)
write_data("Techies",h89.statistic)
write_data("Templar Assasin",h90.statistic)
write_data("Terrorblade",h91.statistic)
write_data("Tidehunter",h92.statistic)
write_data("Timbersaw",h93.statistic)
write_data("Tinker",h94.statistic)
write_data("Tiny",h95.statistic)
write_data("Treant Protector",h96.statistic)
write_data("Troll Warlord",h97.statistic)
write_data("Tusk",h98.statistic)
write_data("Underlord",h99.statistic)
write_data("Undying",h100.statistic)
write_data("Ursa",h101.statistic)
write_data("Vengeful Spirit",h102.statistic)
write_data("Venomancer",h103.statistic)
write_data("Viper",h104.statistic)
write_data("Visage",h105.statistic)
write_data("Warlock",h106.statistic)
write_data("Weaver",h107.statistic)
write_data("Windranger",h108.statistic)
write_data("Winter Wyvern",h109.statistic)
write_data("Witch Doctor",h110.statistic)
write_data("Wraith King",h111.statistic)
write_data("Zeus",h112.statistic)

