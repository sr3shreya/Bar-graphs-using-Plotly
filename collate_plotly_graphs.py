from plotly.subplots import make_subplots
import plotly.graph_objects as go
import pandas as pd
import plotly
import time
import os
from datetime import date

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']),'Desktop','Agent_data_graph')
t_date = str(date.today())
test_path = desktop + '\\' + t_date
if not os.path.exists(test_path):
    os.makedirs(test_path)
    os.makedirs(test_path + '\\Graph_Images')
    os.makedirs(test_path + '\\Graph_HTML')
    #os.makedirs(test_path+'\\Graph_Images\\%Count')
    #os.makedirs(test_path+'\\Graph_Images\\%Coverage')
coverage_list=[]
count_list=[]
def figures_to_html(fig_coverage, fig_count, filename1=test_path+"\\coverage.html", filename2=test_path+"\\count.html"):
    dashboard = open(filename1, 'w')
    dashboard.write("<html><head></head><body>" + "\n")
    for fig in fig_coverage:
        inner_html = fig.to_html().split('<body>')[1].split('</body>')[0]
        dashboard.write(inner_html)
    dashboard.write("</body></html>" + "\n")
    dashboard = open(filename2,'w')
    dashboard.write("<html><head></head><body>" + "\n")
    for fig in fig_count:
        inner_html = fig.to_html().split('<body>')[1].split('</body>')[0]
        dashboard.write(inner_html)
    dashboard.write("</body></html>" + "\n")
