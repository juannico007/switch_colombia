    
import csv
#import pickle
import openpyxl
# from utils.readxlxs import xlxstocsvres

# Historical data wind
dict_fig ={}

with open('dispatch_annual_summary'+'.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    singleData = [[] for x in range(4)]  
    for row in readCSV:
        for col in range(4):
            val = row[col]
            try: 
                val = float(val)
            except ValueError:
                pass
            singleData[col].append(val)

x = [2025,2030,2035,2040,2045,2050]
labels=['Hydro','Minors','Thermal']
import plotly 
import plotly.graph_objs as go

yfuels = [[] for x in range(26)] 
for j in range(6):
    for i in range(26):
        yfuels[i].append(singleData[3][i*6+1+j])

yth=[0]*6
for i in range(24):
    yth=[sum(x)*0.977 for x in zip(yth,yfuels[2+i])]

yfuels2 = [yfuels[0],yfuels[1],yth] 

data=[]
for i in range(3):
#     # Create traces
    trace = go.Bar(
        x = x,
        y = yfuels2[i],
        # mode = 'lines',
        name = labels[i]
    )
    data.append(trace)

layout = go.Layout(
autosize=False,
width=1000,
height=500,
#title='Double Y Axis Example',
yaxis=dict(title=' GWh',
           titlefont=dict(
                   family='Arial, sans-serif',
                    size=18,
                    color='darkgrey'),
            #tickformat = ".0f"
            # exponentformat = "e",
            #showexponent = "none",
            ticks = "inside",
            #range=[20,100]
            ),
# xaxis=dict(range=[axisfixlow,axisfixhig])
)
           
fig = go.Figure(data=data, layout=layout)
dict_fig["aggr"] = plotly.offline.plot(fig, output_type = 'div')

# ###########################################################################
    
from jinja2 import Environment, FileSystemLoader

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("templates/report1.html")

template_vars = {"title" : "Report",
                  "data1": "Each year dispatch",
                  "div_placeholder1A": dict_fig["aggr"]
                  #"div_placeholder1B": dict_fig["string2"],
                  #"div_placeholder1C": dict_fig["string3"],
                  #"div_placeholder1D": dict_fig["string4"],
                  #"div_placeholder1E": dict_fig["string5"],
                  #"data2": "All areas",
                  #"div_placeholder2": graf3,
                  #"data3": ,
                  #"div_placeholder3": ,
                  #"data4": ,
                  #"div_placeholder4": 
                  }

html_out = template.render(template_vars)

Html_file= open("report1.html","w")
Html_file.write(html_out)
Html_file.close()
