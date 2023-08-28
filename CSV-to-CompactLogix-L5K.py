#!/usr/bin/env python
# coding: utf-8

print('Initializing...\n')

import pandas as pd
import os





def print_instructions():
    print('**********************************************')
    print('Instructions:')
    print('1. Export a sheet from Excel with CSV file type')
    print('2. Only use .csv files as an input')
    print('3. Make sure the CSV file includes following columns:')
    print('        BMS Tag Name, Card, Cabinet, Rack, Slot, Channel')
    print('**********************************************')





def get_input():
    status = 'no'
    cols_list = ['IO Tag Name', 'Card', 'Cabinet', 'Rack', 'Slot', 'Channel']
    
    while(status=='no'):

        print('Enter the input file name:')

        input_path = input()
        print('Confirm the path? (y/n):')
        response = input()
        
        if response == 'y':
            status = 'yes'
            
            df_in = pd.read_csv(".\I_O.csv", encoding = "ISO-8859-1", sep=',')
            
#             Check if df_in has all the required columns
            result =  all(elem in df_in.columns for elem in cols_list)
            
            if result:
                df = df_in.loc[:,cols_list]
                print('\nImported ', input_path, '\n')
                
            else:
                print('The Column names do not match with the specified format. Please check the instructions (at the top)')
                status = 'no'
    
    return df





def save_output(df_output, tag, cabinet):
    filename = cabinet + '/' + tag + '.csv'
    df_output.to_csv(filename)
    print('Saved ', filename)    





def make_folder(path):
    try:
        os.mkdir(path)
    except OSError:
        print ("Failed to create %s " % path)
    else:
         print ("Successfully created %s " % path)





def generate_controller_tags(df):
    df_out = pd.DataFrame(columns = df.columns)


        # add addon instruction tag
    df.loc[((df.Card == '1769-IQ16') | (df.Card == '1769-IQ32') | (df.Card == '1756-IB16')),'instruction_tag'] = 'STD_DI'
    df.loc[((df.Card == '1769-OW8') | (df.Card == '1769-OW16') | (df.Card == '1756-OX8I')) | (df.Card == '1756-OW16I'),'instruction_tag'] = 'STD_D2SD'
    df.loc[((df.Card == '1769-IF8') | (df.Card == '1769-IF16') | (df.Card == '1769-IF16C') | (df.Card == '1756-IF8')),'instruction_tag'] = 'STD_AI'


    for cabinet in df.Cabinet.unique():

        if str(cabinet) != 'nan':
            make_folder(cabinet)

            for tag in df.instruction_tag.unique():
                module_tag_list = []

                for item in df.loc[((df.instruction_tag==tag) & (df.Cabinet==cabinet)),:].index:
                    instance = df.loc[item,:].copy()   

                    if ((instance.instruction_tag == 'STD_DI') | (instance.instruction_tag == 'STD_D2SD') | (instance.instruction_tag == 'STD_AI')):
                        slot = int(instance.Slot)
                        channel = int(instance.Channel)
                        instruction_tag = instance.loc['instruction_tag']
                        bms_tag = instance.loc['BMS Tag Name']




                        if ((instance.instruction_tag == 'STD_DI') & (bms_tag != 'SPARE')):
                            instance.loc['controller_tag'] = instruction_tag + ' ' + bms_tag + ' Local:' + str(slot) + ':I.Data.' + str(channel) + ' Local:' + str(slot) + ':I.Fault.' + str(channel)

                        elif ((instance.instruction_tag == 'STD_D2SD') & (bms_tag != 'SPARE')):
                            instance.loc['controller_tag'] = instruction_tag + ' ' + bms_tag + ' ' + bms_tag + '.AutoCmd ' + bms_tag + '.FB0_Input ' + bms_tag + '.FB0_BadIO ' + bms_tag + '.FB1_Input ' + bms_tag + '.FB1_BadIO ' + bms_tag + '.HOA ' + 'Local:' + str(slot) + ':O.Data.' + str(channel) + ' Local:' + str(slot) + ':I.Fault.' + str(channel) + ' ' + bms_tag + '.Output_Inv ' + bms_tag + '.Pulse_Out_0 ' + bms_tag + '.Pulse_Out_0_BadIO ' + bms_tag + '.Pulse_Out_1 ' + bms_tag + '.Pulse_Out_1_BadIO' 

                        elif ((instance.instruction_tag == 'STD_AI') & (bms_tag != 'SPARE')):
                            instance.loc['controller_tag'] = instruction_tag + ' ' + bms_tag + ' Local:' + str(slot) + ':I.Ch' + str(channel) + 'Data' + ' Local:' + str(slot) + ':I.Ch' + str(channel) + 'Status'

                        elif instance.Card == '1769-OF8V':
                            instance.loc['controller_tag'] = 'N/A'
            #                 More complex process for controller tag formation as the 'AO' includes combination of STD_AO and STD_AO_FB.
            #                 Long term solution is to create a column in excel which mentions if feedback  exists


                    else:
                        instance.loc['controller_tag'] = 'N/A'



                    module_tag_list.append(instance)


                df_out = pd.DataFrame(module_tag_list)

                if str(tag) != 'nan':
                    save_output(df_out, tag, cabinet)


if __name__ == "__main__":

    print_instructions()
    df = get_input()
    generate_controller_tags(df)
         
            
    print('***************Finshed******************')

    print('\n\n\nPress enter to close the window')
    temp_var = input()

    

