# -*- coding: utf-8 -*-
"""
Created on Tue Mar  2 08:56:02 2021

@author: VMohanan
"""

# Script for plotting psse channels 

# Import libraries
import os
import numpy as np
import math
import pandas as pd
import plotly.io as pio
pio.renderers.default='browser'
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Import read out file class
from ReadOutFile import outfile



class plot_psse_channels():
    
    def __init__(self, path_to_excel_file, sheet):
        
        self.path_to_input_excel_file = path_to_excel_file
        self.main_sheet = sheet
        

    def read_outfile_selection_data (self):
        
        # Read directory data and plotting data from the input excel file
        self.outfile_directory_data = pd.read_excel(self.path_to_input_excel_file, sheet_name = self.main_sheet)
        
        self.overlayplot_selection = int(self.outfile_directory_data['Overlay multiple files'][0])
        self.number_of_row = int(self.outfile_directory_data['No of Graph Rows'][0])
        self.number_of_column = int(self.outfile_directory_data['No of Graph Colums'][0])
        

    def read_channel_data_sheet(self, sheetname):
        
        self.channel_data_dict = dict()
        
        channel_selection_data = pd.read_excel(self.path_to_input_excel_file, sheet_name = sheetname)
        
        # Skip rows channels which are not selected
        channel_selection_data = channel_selection_data.loc[channel_selection_data['Channel Selection'] == 1]
        
        
        # Creating a dictionary to store the selected channels with key as graph number
        for row in range(len(channel_selection_data)):            
            
            graph_number = channel_selection_data.iloc[row]['Graph Number Clock Wise']
            
            if self.channel_data_dict.has_key(graph_number):
                self.channel_data_dict[graph_number].append(channel_selection_data.iloc[row,:4].tolist())
            else:
                self.channel_data_dict[graph_number] = [channel_selection_data.iloc[row,:4].tolist()]    

    def read_out_file(self, path_to_outfile):    
        
        try:
            results = outfile(path_to_outfile)
            self.ChannelList = results.chans
            self.TimeData = np.array(results.time)
            self.PSSEData = np.array(results.data)
        except:
            print('\nInvalid outfile path or name')
            print('\nSkipping case{}'.format(path_to_outfile))
            return -1
    

    def add_plot_traces(self, fig, graph_number, outfile_description, overlay_selection = 0):
        
        for channel_data in self.channel_data_dict[graph_number]:
            
            channel_name, plot_title, multiplication_factor, addition_factor = channel_data[0:4]
            if math.isnan(multiplication_factor) or math.isnan(addition_factor):
                    
                    multiplication_factor = 1
                    addition_factor = 0
                    Ylabel = 'pu'
                
                    if 'PELEC:' in channel_name or ('POWR' in channel_name and 'TO' not in channel_name):
                        multiplication_factor = 100.0
                        Ylabel = 'MW'
                        
                    if 'QELEC:' in channel_name:
                        multiplication_factor = 100.0
                        Ylabel = 'MVar'
                        
                    if 'FREQ' in channel_name:
                        multiplication_factor = 50.0
                        addition_factor = 50.0
                        Ylabel = 'Hz'
                        
                    if 'P FLOW' in channel_name:
                        Ylabel = 'MW'
                        
                    if 'Q FLOW' in channel_name:
                        Ylabel = 'MVAr'
                        
                    if 'POWR' in channel_name and 'TO' in channel_name:
                        Ylabel = 'MW'
  
            try:
                y_data = self.PSSEData[self.ChannelList.index(channel_name)]*multiplication_factor + addition_factor
            except:
                print('\nInvalid Channel Name {} or Channel does not exist'.format(channel_name))
                print('\nSkipping this channel')
                continue
            Xlabel = 'time(s)'
            
            
            if graph_number > self.number_of_row * self.number_of_column:
                print('\nGraph number exceeding maximum graph possible')
                print('\nSkipping channel {}'.format(channel_name))
                continue
        
            
            if overlay_selection == 1:
                legend = outfile_description + '_' + Ylabel
                subplot_title = channel_name
            elif math.isnan(plot_title):
                legend = channel_name
                subplot_title = 'Graph-{}'.format(graph_number)
            else:
                legend = channel_name
                subplot_title = plot_title
            
            
            # Identifying the graph row and column from the graph number.
            count = 1
            break_loop = 0
            for Row in range(1, self.number_of_row + 1):
                for Col in range(1, self.number_of_column + 1):
                    if count == graph_number:
                        break_loop = 1
                        break
                    count += 1
                if break_loop == 1:
                    break
            
            fig.add_trace(go.Scatter(x = self.TimeData,  y = y_data, name = legend, showlegend = True), row = Row, col = Col)
            fig.update_yaxes(title_text = Ylabel, row = Row, col = Col)
            fig.update_xaxes(title_text = Xlabel, row = Row, col = Col)
            fig.layout.annotations[graph_number-1].update(text = subplot_title)
    
    def process_selection(self):
        
        if self.overlayplot_selection == 1:
            # If overlay plot is selected, the script will not check any other selection
            
            plot.read_channel_data_sheet('OverlayPlotData')
            
            self.outfile_directory_data = self.outfile_directory_data.loc[self.outfile_directory_data['File Selection'] == 1]
            
            fig = make_subplots(rows =  self.number_of_row, cols = self.number_of_column, subplot_titles = ['PlaceHolder']*self.number_of_row * self.number_of_column)
            print('\nProcessing Overlay plots..')
            
            for graph_number in self.channel_data_dict.keys():
                directory_data = zip(self.outfile_directory_data['PSSE Out File Directory'], self.outfile_directory_data['PSSE Out File Name'], self.outfile_directory_data['Description'])
                
                print('\nProcessing data for graph {}..'.format(graph_number))
                
                
                for outfile_path, outfile_name, file_description in directory_data:
                    path_to_outfile =  os.path.join(outfile_path, outfile_name) 
                    read_flag = self.read_out_file(path_to_outfile)
                    
                    if read_flag == -1:
                        continue
                    
                    self.add_plot_traces(fig, graph_number, file_description, 1)
                    
            #Creating output directory
            print('\nSaving plots..')
            output_directory = os.path.join(outfile_path, 'Results')
            results_directory = os.path.join(output_directory, 'OverlayPlots')
            try:
                os.makedirs(results_directory)
            except:
                pass
            fig.write_html(results_directory + '/Overlay.html')
            print('\nFinished plotting')   
        
        # If overlay option is not selected, the script will check the channel output to excel selection option. 
        else:
            outfile_to_excel_selection = max(self.outfile_directory_data['Channel Output Selection'])
            
            # If channel output to excel selection option 
            if outfile_to_excel_selection == 1:
                
                self.outfile_directory_data = self.outfile_directory_data.loc[self.outfile_directory_data['Channel Output Selection'] == 1]
                
                print('\nGenerating excelfile with channel list from outfiles') 
                directory_data = zip(self.outfile_directory_data['PSSE Out File Directory'], self.outfile_directory_data['PSSE Out File Name'])
                
                for outfile_path, outfile_name in directory_data:
                    print('\nProcessing outfile: {}'.format(outfile_name.split('.')[0]))
                    path_to_outfile =  os.path.join(outfile_path, outfile_name) 
                    read_flag = self.read_out_file(path_to_outfile)
                    
                    col_length = len(self.ChannelList)
                    channelData = pd.DataFrame({'Channel Name': self.ChannelList,
                                                'Plot title' : ['N/A']*col_length,
                                                'MultiplicationFactor' : ['N/A']*col_length,
                                                'AdditionFactor' : ['N/A']*col_length,
                                                'Graph Number Clock Wise' : [1]*col_length,
                                                'Channel Selection' : [1]*col_length,
                                                })
                    
                    # Save the files to the output directory
                    output_directory = os.path.join(outfile_path, 'Channel_List')
                    try:
                        os.makedirs(output_directory)
                    except:
                        pass
                    excel_file_name = outfile_name.split('.')[0]
                    writer = pd.ExcelWriter(output_directory + '/' + excel_file_name+'.xlsx', engine='xlsxwriter')
                    channelData.to_excel(writer, sheet_name='Sheet1', index=False, columns = ['Channel Name', 'Plot title', 
                                                                                              'MultiplicationFactor', 
                                                                                              'AdditionFactor',
                                                                                              'Graph Number Clock Wise',
                                                                                              'Channel Selection'])
                    writer.save()
                    
            else:
                print('\nProcessing selected individual outfiles')
                
                self.outfile_directory_data = self.outfile_directory_data.loc[self.outfile_directory_data['File Selection'] == 1]
                
                directory_data = zip(self.outfile_directory_data['PSSE Out File Directory'],
                                     self.outfile_directory_data['PSSE Out File Name'],
                                     self.outfile_directory_data['Description'],
                                     self.outfile_directory_data['Channel Data sheet name'])
                
                #fig = make_subplots(rows = self.number_of_row, cols = self.number_of_column, subplot_titles = ['dummy']*self.number_of_row * self.number_of_column)
                
                for outfile_path, outfile_name, file_description, sheet_name in directory_data:
                    
                    fig = make_subplots(rows = self.number_of_row, cols = self.number_of_column, subplot_titles = ['PlaceHolder']*self.number_of_row * self.number_of_column)
                    
                    print('\nProcessing outfile: {}'.format(outfile_name.split('.')[0]))
                    
                    path_to_outfile =  os.path.join(outfile_path, outfile_name) 
                    read_flag = self.read_out_file(path_to_outfile)
                    
                    plot.read_channel_data_sheet(sheet_name)
                    
                    if read_flag == -1:
                        continue
                    
                    for graph_number in self.channel_data_dict.keys():
                        
                        self.add_plot_traces(fig, graph_number, file_description, 0)
                    
                    #Creating output directory
                    print('\nSaving plots..')
                    output_directory = os.path.join(outfile_path, 'Results')
                    fileName = outfile_name.split('.')[0]
                    results_directory = os.path.join(output_directory, fileName)
                    try:
                        os.makedirs(results_directory)
                    except:
                        pass
                    fig.write_html(results_directory + '/'+fileName+'.html')
                    print('\nFinished plotting')   
                    del fig 
                    
 
plot = plot_psse_channels('PlotInputData.xlsx', 'Data')
plot.read_outfile_selection_data()
plot.process_selection()