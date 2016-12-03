"""Plot Graph by xlrd module"""
import xlrd
file = xlrd.open_workbook('Data.xls')
my_data_advantage = file.sheet_by_name('advantage')
my_data_winrate = file.sheet_by_name('winRate')
my_data_matchs = file.sheet_by_name('matchs')

def collect_data(number):
    """Collects data from heroes A-Z"""
    col, row = 0, 0
    hero_advantage = []
    hero_name = []
    #number = 1 is abaddon 112 is Zeus
    for i in range(1,113): #Data abaddon start with 1,0 1-13 character
        hero_advantage.append(my_data_advantage.cell(number,i).value)
    for i in range(1,113):
        hero_name.append(my_data_advantage.cell(i,0).value)
    return hero_advantage, hero_name

def plot():
    """Plot Graph by pylab"""
    import pylab as plt
    hero = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20,\
            21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38,\
            39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56,\
            57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74,\
            75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92,\
            93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108,\
            109, 110, 111, 112] #index of hero
    LABELS, names = collect_data(1)
    plt.bar(hero, LABELS, align='center')
    plt.xticks(range(112), names) # (x width, names)
    plt.gca().set_xlim(0.1, 30)
    plt.show()

def plot_plotly():
    """"""
    import plotly.plotly as py
    import plotly.graph_objs as go
    hero_advantage = []
    hero_name = []
    for i in range(1, 3):
        hero_advantage, hero_name = collect_data(i)
        trace = go.Bar(
            x=hero_advantage,
            y=hero_name
            )
        data = [trace]
        py.iplot(data, filename=hero_name[i-1])

plot_plotly()
