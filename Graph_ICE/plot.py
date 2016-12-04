"""Plot Graph by xlrd module"""
import xlrd
file = xlrd.open_workbook('Data.xls')
my_data_advantage = file.sheet_by_name('advantage')
my_data_winrate = file.sheet_by_name('winRate')
my_data_matchs = file.sheet_by_name('matchs')

def collect_data(number):
    """Collects data from heroes A-Z"""
    hero_advantage, hero_name, hero_winrate, hero_match = [], [], [], []
    #number = 1 is abaddon 112 is Zeus
    for i in range(1,113): #Data abaddon start with 1,0 1-13 character
        hero_advantage.append(my_data_advantage.cell(number,i).value)
        hero_name.append(my_data_advantage.cell(i,0).value)
        hero_winrate.append(my_data_winrate.cell(number,i).value)
        hero_match.append(my_data_matchs.cell(number,i).value)
    return hero_advantage, hero_name, hero_winrate, hero_match

def plot_comparison_graphs():
    """Show graphs of 112 Heroes sort by Heroes Name"""
    import plotly.offline as offline
    from plotly.offline import download_plotlyjs, plot
    from plotly.graph_objs import Bar
    hero_advantage, hero_name, hero_winrate, hero_match = [], [], [], []
    for i in range(1, 113):
        hero_advantage, hero_name, hero_winrate, hero_match = collect_data(i)
        Advantage = Bar(
            {"x":hero_name,
            "y":hero_advantage,"type":"bar","name":'Advantage'
            }
            )
        WinRate = Bar(
            {"x":hero_name,
            "y":hero_winrate,"type":"bar","name":'WinRate'
            }
            )
        MatchPlay = Bar(
            {"x":hero_name,
            "y":hero_match,"type":"bar","name":'MatchPlayed'
            })
        data = [Advantage, WinRate, MatchPlay]
        offline.plot({'data':data, 'layout':{'title': hero_name[i-1]+" - Advantage is good for range > 0 Disadvantage otherwise", 'font': dict(size=16)}}, filename=str(hero_name[i-1])+".html", image='png')

plot_comparison_graphs()
