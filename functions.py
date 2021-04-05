###!/usr/bin/env python
# coding: utf-8

### Import Packages
import pandas as pd
import ixmp
import message_ix
import xlsxwriter
import os
import datetime
import openpyxl

from message_ix.utils import make_df
from message_ix.reporting import Reporter
from openpyxl import load_workbook


# Define Helper Functions

### Spreadsheet Functions

def make_filepath(name): 
    # Create filename and check for duplicates
    directory = os.getcwd()
    folder = directory + '/Data Sheets' # store sheets in subfolder. Probably need to autocreate this folder.
    count = 1
    while (os.path.exists(folder+'/'+name+'.xlsx') == True):
        name = name.split('_')[0] + '_' + str(count)
        count = count + 1
    filepath = folder + '/' + name + '.xlsx'
    return filepath, name

def write_file(file, df, df_type):
    
    # Open writer
    book = load_workbook(file)
    writer = pd.ExcelWriter(file, engine='openpyxl')
    writer.book = book
    
    # Write df to specified spreadsheet
    df.to_excel(writer, df_type)
    
    # Save file
    writer.save()


# ### Run Baseline Model

def run_baseline():
    
    # Create scenario
    mp = ixmp.Platform()
    scenario = message_ix.Scenario(mp, model='Westeros Electrified', scenario='input_baseline', version='new')

    # Create timeline
    history = [300]
    model_horizon = [310, 320, 330, 340, 350]
    scenario.add_horizon(
        year=history + model_horizon,
        firstmodelyear=model_horizon[0]
    )

    # Define regions
    country = 'Westeros'
    scenario.add_spatial_sets({'country': country})

    # Define commodities
    scenario.add_set("commodity", ["electricity", "light"])
    scenario.add_set("level", ["secondary", "final", "useful"])

    # Add technology sets
    scenario.add_set("technology", ['coal_ppl', 'wind_ppl', 'pv_ppl', 'grid', 'bulb'])
    #scenario.add_set("technology", ['coal_ppl', 'wind_ppl', 'pv_ppl', 'grid', 'bulb', 'battery'])
    #scenario.add_set("technology", ['coal_ppl', 'wind_ppl', 'grid', 'bulb'])
    scenario.add_set("mode", "standard")

    # # Add demand
    # gdp_profile = pd.Series([1., 1.5, 1.9], index=pd.Index(model_horizon, name='Time'))
    # demand_per_year = 40 * 12 * 1000 / 8760
    # light_demand = pd.DataFrame({
    #         'node': country,
    #         'commodity': 'light',
    #         'level': 'useful',
    #         'year': model_horizon,
    #         'time': 'year',
    #         'value': (100 * gdp_profile).round(),
    #         #'value': demand_per_year,
    #         'unit': 'GWa',
    #     })
    # scenario.add_par("demand", light_demand)

    # Create input/output connections
    year_df = scenario.vintage_and_active_years()
    vintage_years, act_years = year_df['year_vtg'], year_df['year_act']
    unit = 'MWa'
    mp.add_unit('MWa')

    base = {
        'node_loc': country,
        'year_vtg': vintage_years,
        'year_act': act_years,
        'mode': 'standard',
        'time': 'year',
        'unit': '-',
    }

    base_input = make_df(base, node_origin=country, time_origin='year')
    base_output = make_df(base, node_dest=country, time_dest='year')

    # Light input/output
    bulb_out = make_df(base_output, technology='bulb', commodity='light', 
                       level='useful', value=1.0)
    scenario.add_par('output', bulb_out)

    bulb_in = make_df(base_input, technology='bulb', commodity='electricity',  
                      level='final', value=1.0)
    scenario.add_par('input', bulb_in)

    # Grid input/output
    grid_efficiency = 0.9
    grid_out = make_df(base_output, technology='grid', commodity='electricity', 
                       level='final', value=grid_efficiency)
    scenario.add_par('output', grid_out)

    grid_in = make_df(base_input, technology='grid', commodity='electricity',
                      level='secondary', value=1.0)
    scenario.add_par('input', grid_in)

    # Power plants
    coal_out = make_df(base_output, technology='coal_ppl', commodity='electricity', 
                       level='secondary', value=1., unit=unit)
    scenario.add_par('output', coal_out)

    wind_out = make_df(base_output, technology='wind_ppl', commodity='electricity', 
                       level='secondary', value=1., unit=unit)
    scenario.add_par('output', wind_out)

    pv_out = make_df(base_output, technology='pv_ppl',commodity='electricity', 
                        level='final', value=1., unit=unit)
    scenario.add_par('output', pv_out)

    # # Battery
    # battery_in = make_df(base_input, technology='battery', commodity='electricity',
    #                     level='final', value=1, unit=unit)
    # scenario.add_par('input', battery_in)

    # battery_out = make_df(base_output, technology='battery', commodity='electricity', 
    #                      level='final', value=1, unit=unit)
    # scenario.add_par('output', battery_out)

    # Capacity factors 
    base_capacity_factor = {
        'node_loc': country,
        'year_vtg': vintage_years,
        'year_act': act_years,
        'time': 'year',
        'unit': '-',
    }
    capacity_factor = {
        'coal_ppl': 1,
        'wind_ppl': 0.2,
        'pv_ppl': 0.15,
        'bulb': 1, 
        #'battery': 1
    }

    for tec, val in capacity_factor.items():
        df = make_df(base_capacity_factor, technology=tec, value=val)
        scenario.add_par('capacity_factor', df)

    # Emissions
    mp.add_unit('tCO2/kWa')
    scenario.add_set('emission', 'CO2')
    scenario.add_cat('emission', 'GHG', 'CO2')
    base_emission_factor = {
        'node_loc': country,
        'year_vtg': vintage_years,
        'year_act': act_years,
        'mode': 'standard',
        'unit': 'tCO2/kWa',
    }
    emission_factor = make_df(base_emission_factor, technology= 'coal_ppl', emission= 'CO2', value = 7.48)
    scenario.add_par('emission_factor', emission_factor)

    # Tech lifetimes
    base_technical_lifetime = {
        'node_loc': country,
        'year_vtg': model_horizon,
        'unit': 'y',
    }
    lifetime = {
        'coal_ppl': 40,
        'wind_ppl': 20,
        'pv_ppl': 20,
        'bulb': 1,
    }

    for tec, val in lifetime.items():
        df = make_df(base_technical_lifetime, technology=tec, value=val)
        scenario.add_par('technical_lifetime', df)

    # # Set tech growth 
    # allowed_growth = 0.50
    # base_growth = {
    #     'node_loc': country,
    #     'year_act': model_horizon,
    #     'time': 'year',
    #     'unit': '-',
    # }
    # growth_technologies = [
    #     "coal_ppl", 
    #     "wind_ppl",
    #     "pv_ppl"
    # ]

    # for tec in growth_technologies:
    #     df = make_df(base_growth, technology=tec, value=allowed_growth) 
    #     scenario.add_par('growth_activity_up', df)

    # Add objective function
    scenario.add_par("interestrate", model_horizon, value=0.05, unit='-')
    
    base_inv_cost = {
        'node_loc': country,
        'year_vtg': model_horizon,
        'unit': 'USD/kW',
    }

    mp.add_unit('USD/kW') 
    mp.add_unit('kWa')

    costs = {
        'coal_ppl': 1500,
        'wind_ppl': 1100,
        'pv_ppl': 4000,
        'bulb': 5,
    } # in $ / kW (specific investment cost)

    for tec, val in costs.items():
        df = make_df(base_inv_cost, technology=tec, value=val)
        scenario.add_par('inv_cost', df)

    base_fix_cost = {
        'node_loc': country,
        'year_vtg': vintage_years,
        'year_act': act_years,
        'unit': 'USD/kWa',
    }

    # in $ / kW / year (every year a fixed quantity is destinated to cover part of the O&M costs
    # based on the size of the plant, e.g. lightning, labor, scheduled maintenance, etc.)

    costs = {
        'coal_ppl': 40,
        'wind_ppl': 40,
        'pv_ppl': 25
    }

    for tec, val in costs.items():
        df = make_df(base_fix_cost, technology=tec, value=val)
        scenario.add_par('fix_cost', df)

    base_var_cost = {
        'node_loc': country,
        'year_vtg': vintage_years,
        'year_act': act_years,
        'mode': 'standard',
        'time': 'year',
        'unit': 'USD/kWa',
    }

    # in $ / kWa (costs associatied to the degradation of equipment when the plant is functioning
    # per unit of energy produced kW·year = 8760 kWh.
    # Therefore this costs represents USD per 8760 kWh of energy). Do not confuse with fixed O&M units.

    costs = {
        'coal_ppl': 24.4,
        'grid': 47.8,
    }
    
    for tec, val in costs.items():
        df = make_df(base_var_cost, technology=tec, value=val)
        scenario.add_par('var_cost', df)

    # # peak_load_factor(node,commodity,level,year,time)
    # peak_load_factor = pd.DataFrame({
    #         'node': country,
    #         'commodity': 'electricity',
    #         'level' : 'secondary',       
    #         'year': model_horizon,
    #         'time' : 'year',
    #         'value' : 2,
    #         'unit' : '-'})

    # scenario.add_par('peak_load_factor', peak_load_factor)

    # base_reliability = pd.DataFrame({
    #         'node': country,
    #         'commodity': 'electricity',
    #         'level' : 'secondary', 
    #         'unit': '-',
    #         'time': 'year',
    #         'year_act': model_horizon})

    # # adding wind ratings to the respective set 
    # scenario.add_set('rating', ['r1', 'r2'])

    # # adding rating bins for wind power plant
    # rating_bin = make_df(base_reliability, technology= 'wind_ppl', value = 0.1, rating= 'r1')
    # scenario.add_par('rating_bin', rating_bin)

    # rating_bin = make_df(base_reliability, technology= 'wind_ppl', value = 0.9, rating= 'r2')
    # scenario.add_par('rating_bin', rating_bin)

    # # adding reliability factor for each rating of wind power plant
    # reliability_factor = make_df(base_reliability, technology= 'wind_ppl', value = 0.8, rating= 'r1')
    # scenario.add_par('reliability_factor', reliability_factor)

    # reliability_factor = make_df(base_reliability, technology= 'wind_ppl', value = 0.05, rating= 'r2')
    # scenario.add_par('reliability_factor', reliability_factor)

    # # considering coal power plant as firm capacity (adding a reliability factor of 1)
    # reliability_factor = make_df(base_reliability, technology= 'coal_ppl', value = 1, rating= 'firm')
    # scenario.add_par('reliability_factor', reliability_factor)


    scenario.commit('Baseline')
    #fp = os.getcwd() + '/Data Sheets/Template Scenario.xlsx'
    #scenario.to_excel(fp)

    # Close model 
    mp.close_db()


### Run New Model

def run_model_from_sheet(filepath, scen_name):
    unit = 'MWa'
    model_horizon = [310, 320, 330, 340, 350]
    country = 'Westeros'
    history = [300]
    grid_efficiency = 0.9
    #emi_bound = 500
    # cost_bound = 100000
    #renewable_min = 0.5
    
    # Read inputs sheet
    xlsx = pd.ExcelFile(filepath)
    #cap_df = xlsx.parse('Cap Inputs')
    demand_df = xlsx.parse('Population Inputs')
    #storage_df = xlsx.parse('Storage Inputs')
    
    # Open model platform
    mp = ixmp.Platform()
    
    # Clone baseline scenario
    time = datetime.datetime.now()
    model = 'Westeros Electrified'
    base = message_ix.Scenario(mp, model=model, scenario='input_baseline')
    scen = base.clone(model, scen_name, str(time), keep_solution=False)
    scen.check_out()
    
    # Add demand
    pop_list = []
    for index, row in demand_df.iterrows():
        pop_list.append(row[0])
    #print(demand_list)
    demand_list = [i*1000/8760/1000 for i in pop_list] #people x 1000kWh/person / # hours in a year / 1000 to get MWa
    
    demand_input = pd.Series(demand_list, index=pd.Index(model_horizon, name='Time'))
    light_demand = pd.DataFrame({
            'node': country,
            'commodity': 'light',
            'level': 'useful',
            'year': model_horizon,
            'time': 'year',
            'value': demand_input,
            'unit': unit,
        })
    scen.add_par("demand", light_demand)
    
    # # Add cost bound
    # scen.add_par('total_cost', [country, 'all'], value=cost_bound, unit='USD')

    # Add historical activity
    historic_demand = 0.85 * demand_list[0]
    historic_generation = historic_demand / grid_efficiency

    base_activity = {
        'node_loc': country,
        'year_act': history,
        'mode': 'standard',
        'time': 'year',
        'unit': unit,
    }
    
    old_activity = {
        'coal_ppl': 1 * historic_generation,
        'wind_ppl': 0 * historic_generation,
        'pv_ppl': 0 * historic_generation
    }

    for tec, val in old_activity.items():
        df = make_df(base_activity, technology=tec, value=val)
        scen.add_par('historical_activity', df)   
        
    # Add base capacities
    capacity_factor = {
        'coal_ppl': 1,
        'wind_ppl': 0.2,
        'pv_ppl': 0.15,
        'bulb': 1, 
        #'battery': 1
    }
    
    act_to_cap = {
        'coal_ppl': 1 / 10 / capacity_factor['coal_ppl'] / 2, # 20 year lifetime
        'wind_ppl': 1 / 10 / capacity_factor['wind_ppl'] / 2,
        'pv_ppl': 1 / 10 / capacity_factor['pv_ppl']/ 2
    }
    
    base_capacity = {
        'node_loc': country,
        'year_vtg': history,
        'unit': unit,
    }


    for tec in act_to_cap:
        value = old_activity[tec] * act_to_cap[tec]
        df = make_df(base_capacity, technology=tec, value=value)
        scen.add_par('historical_new_capacity', df)
        
    # Add activity lower bounds
    # coal_percent = cap_df.loc[cap_df['Technology'] == 'coal_ppl', 'Capacity'].iloc[0]
    # wind_percent = cap_df.loc[cap_df['Technology'] == 'wind_ppl', 'Capacity'].iloc[0]
    # pv_percent = cap_df.loc[cap_df['Technology'] == 'pv_ppl', 'Capacity'].iloc[0]
   
    # Total energy share
    # share_coal = 'share_coal'
    # share_wind = 'share_wind'
    # share_pv = 'share_pv'
    # scen.add_set('shares', share_coal)
    # scen.add_set('shares', share_wind)
    # scen.add_set('shares', share_pv)
    
        
    # Add emission bound
    scen.commit('Solving BAU')
    scen.solve()
    rep = Reporter.from_scenario(scen)
    emi_key = rep.full_key('emi').drop('h', 'yv')
    act_key = rep.full_key('ACT').drop('h', 'yv')
    emi = rep.get(emi_key).to_dataframe()
    act = rep.get(act_key).to_dataframe()
    bau_emi = emi.to_numpy().sum()
    scen.remove_solution()
    
    #historic_emi = old_activity['coal_ppl']*7.4*30
    #emi_bound = xlsx.parse('Emission Bound').iloc[0]['Emission Bound']/100*historic_emi
    emi_bound = xlsx.parse('Emission Bound').iloc[0]['Emission Bound']/100 * bau_emi / 5
    
    #DEBUG
    emi_df = pd.DataFrame(data={'Emission Limit': [emi_bound], 'BAU': [bau_emi]})
    write_file(filepath, emi_df, 'Emission Limit')
    write_file(filepath, emi, 'BAU Emissions')
    write_file(filepath, act, 'BAU Activity')
    
    scen.check_out()
    scen.add_par('bound_emission', [country, 'GHG', 'all', 'cumulative'], value = emi_bound, unit='MtCO2')

    # Add renewable energy shares
    wind_max = xlsx.parse('Wind Percent').iloc[0]['Wind Percent']/100
    shares = 'share_wind_electricity'
    scen.add_set('shares', shares)
    
    # Define renewable share
    type_tec = 'electricity_renewable'
    scen.add_cat('technology', type_tec, 'wind_ppl')
    scen.add_cat('technology', type_tec, 'pv_ppl')
    df = pd.DataFrame({'shares': [shares],
                   'node_share': country,
                   'node': country,
                   'type_tec': type_tec,
                   'mode': 'standard',
                   'commodity': 'electricity',
                   'level': 'secondary',
    })
    scen.add_set('map_shares_commodity_total', df)
    
    # Define wind share (of renewable)
    type_tec = 'electricity_wind'
    scen.add_cat('technology', type_tec, 'wind_ppl')
    df = pd.DataFrame({'shares': [shares],
                   'node_share': country,
                   'node': country,
                   'type_tec': type_tec,
                   'mode': 'standard',
                   'commodity': 'electricity',
                   'level': 'secondary',
        })
    scen.add_set('map_shares_commodity_share', df)
    
    # Set as upper bound
    df = pd.DataFrame({'shares': shares,
                   'node_share': country,
                   'year_act': [310],
                   'time': 'year',
                   'value': [wind_max],
                   'unit': '-'})
    scen.add_par('share_commodity_up', df)
    
    df = pd.DataFrame({'shares': shares,
                   'node_share': country,
                   'year_act': [320],
                   'time': 'year',
                   'value': [wind_max],
                   'unit': '-'})
    scen.add_par('share_commodity_up', df)
    
    df = pd.DataFrame({'shares': shares,
                   'node_share': country,
                   'year_act': [330],
                   'time': 'year',
                   'value': [wind_max],
                   'unit': '-'})
    scen.add_par('share_commodity_up', df)
    
    df = pd.DataFrame({'shares': shares,
                   'node_share': country,
                   'year_act': [340],
                   'time': 'year',
                   'value': [wind_max],
                   'unit': '-'})
    scen.add_par('share_commodity_up', df)
    
    df = pd.DataFrame({'shares': shares,
                   'node_share': country,
                   'year_act': [350],
                   'time': 'year',
                   'value': [wind_max],
                   'unit': '-'})
    scen.add_par('share_commodity_up', df)
    
    # Solve scenario
    #scen.to_excel(os.getcwd() + '/Data Sheets/' + scen_name + ' Parameters.xlsx')
    scen.commit('Solving ' + scen_name)
    
    # DEBUG SECTION
    scen.solve()
        
    # Save Results
    save_results(scen, filepath)
    
    mp.close_db()
    
    # Catch infeasibility errors
#     try:
#         scen.solve()
        
#         # Save Results
#         save_results(scen, filepath)
#         #make_plots(scen, filepath)
        
#     except: 
#         print("Problem is infeasible.")
#     finally:
#         mp.close_db()


# ### Link to Interface and Save Results

# In[ ]:


def save_results(scen, fp):
    rep = Reporter.from_scenario(scen)
    #rep.set_filters(t=['coal_ppl', 'wind_ppl'])
    rep.set_filters(t=['coal_ppl', 'wind_ppl', 'pv_ppl'])
    #rep.set_filters(t=['coal_ppl', 'wind_ppl', 'pv_ppl', 'battery'])
    
    #to_get = ['CAP', 'CAP_NEW', 'ACT', 'emi']
    #to_get = ['ACT', 'CAP', 'emi']
    #to_get = ['inv_cost', 'fix_cost', 'var_cost', 'CAP', 'emi']
    #to_get = ['inv_cost', 'fix_cost', 'var_cost', 'emi', 'CAP', 'ACT']
    to_get = ['fom', 'vom', 'emi', 'CAP', 'ACT']
    key_dict = {}
    for prop in to_get: 
        key = rep.full_key(prop).drop('h', 'yv')
        key_df = rep.get(key).to_dataframe()
        write_file(fp, key_df, prop)
        key_dict[prop]=key_df
    
    to_get = ['inv', 'CAP_NEW']
    for prop in to_get:
        key = rep.full_key(prop)
        key_df = rep.get(key).to_dataframe()
        write_file(fp, key_df, prop)
        key_dict[prop]=key_df
        
    # Get Cost
    # costs={}
    # inv_key = rep.full_key('inv').drop('h', 'yv')
    # fix_key = rep.full_key('fom').drop('h', 'yv')
    # var_key = rep.full_key('vom').drop('h', 'yv')
    # costs['Inv Cost']=rep.get(inv_key).to_dataframe()
    # costs['Fix Cost']=rep.get(fix_key).to_dataframe()
    # costs['Var Cost']=rep.get(var_key).to_dataframe()
    # total = costs['Inv Cost'].to_numpy().sum() + costs['Fix Cost'].to_numpy().sum() + costs['Var Cost'].to_numpy().sum()
    # total_cost = pd.DataFrame(data={'Cost':[total]})
    total_cost = pd.DataFrame(data={'Cost':[scen.var('OBJ')['lvl']]})
    write_file(fp, total_cost, 'Total Cost')
    
    # Get Emissions
    emi_key = rep.full_key('emi').drop('h', 'yv')
    emi = rep.get(emi_key).to_dataframe()
    total_emi = pd.DataFrame(data={'Emissions':[emi.to_numpy().sum()]})
    write_file(fp, total_emi, 'Total Emissions')

def process_inputs(app):
    # Grab input values
    name = app._name.value
    pop1 = app._demand.value[0]
    pop2 = app._demand.value[1]
    pop_dif = pop2-pop1
    input_demand = [pop1, pop1+pop_dif/4, pop1 + pop_dif/2, pop1 + 3*pop_dif/4, pop2]
    # coal = app._coal.value
    # pv = app._solar.value
    # wind = app._wind.value
    #storage = app._storage.value
    
    # Create destination filepath and get updated name
    fp, scen_name = make_filepath(name)

    # Save capacity inputs as dataframe
    # input_cap = {
    #     'Technology':['coal_ppl', 'pv_ppl', 'wind_ppl'], 
    #     'Capacity':[coal, pv, wind]
    # }
    # cap_df = pd.DataFrame(data=input_cap)
    
    # # Save storage input as dataframe
    # input_storage = {
    #     'Technology': ['battery'],
    #     'Capacity': [storage]
    # }
    # storage_df = pd.DataFrame(data=input_storage)
    
    # Save demand inputs as dataframe
    demand_df = pd.DataFrame(data=input_demand)
    
    # Create spreadsheet for scenario
    writer = pd.ExcelWriter(fp, engine='openpyxl')
    #cap_df.to_excel(writer, index=False, sheet_name="Cap Inputs")
    demand_df.to_excel(writer, index=False, sheet_name="Population Inputs")
    #storage_df.to_excel(writer, index=False, sheet_name='Storage Inputs')
    writer.save()
    writer.close()
    
    # Write capacities to spreadsheet - Use this to add additional sheets
    # write_file(fp, df, 'Capacity')
    df = pd.DataFrame(data={'Emission Bound':[app._emibound.value]})
    write_file(fp, df, 'Emission Bound')
    
    df = pd.DataFrame(data={'Wind Percent':[app._wind.value]})
    write_file(fp, df, 'Wind Percent')
    
    # Rerun model from spreadsheet
    run_model_from_sheet(fp, scen_name)

    # Read from saved spreadsheet
    xlsx = pd.ExcelFile(fp)
    #cap_sheet = xlsx.parse('CAP')
    #to_get = ['CAP', 'CAP_NEW', 'ACT', 'emi']
    #to_get = ['ACT', 'CAP', 'emi']
    #to_get = ['inv_cost', 'fix_cost', 'var_cost', 'CAP', 'emi']
    #to_get = ['inv_cost', 'fix_cost', 'var_cost', 'emi', 'CAP', 'ACT']
    to_get = ['inv', 'fom', 'vom', 'emi', 'CAP', 'ACT', 'CAP_NEW']
    sheet_dict = {}
    for prop in to_get: 
        sheet = xlsx.parse(prop)
        sheet_dict[prop]=sheet
    #print(cap_sheet)
    
    cost = xlsx.parse('Total Cost').iloc[0]['Cost']
    emissions = xlsx.parse('Total Emissions').iloc[0]['Emissions']
    
    return cost, emissions, sheet_dict, scen_name