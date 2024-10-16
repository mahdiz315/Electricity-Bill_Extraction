import streamlit as st
import pdfplumber
import re
import pandas as pd
import os
import numpy as np
import xlsxwriter
from openpyxl.styles import PatternFill
from openpyxl.worksheet.dimensions import SheetFormatProperties
from datetime import datetime
import datetime as dt 

# Function to extract data from a PDF file
def extract_data(pdf_file):
        # Initialize a dictionary with 0 values for all keywords
    extracted_data = {keyword: 0 for keyword in keywords}

    with pdfplumber.open(pdf_file) as pdf:
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]
            text = page.extract_text()
            text = text.replace("Euel", "Fuel")

            # Replace "FEb" with "EEb"
            text = text.replace("Eeb", "Feb")
            #text = text.replace("Regulatoiy fee", "Regulatory fee")
            

            # Initialize flags for Rate and Constant
            rate_flag = False
            constant_flag = False
            service_to_flag = False

            for line in text.split('\n'):
                # Scenario 4: Find exact word "Rate"
                if "Rate:" in line:
                    rate_flag = True
                    # Use regex to extract 2 words after "Rate"
                    rate_match = re.search(r'Rate:\s*(\S+\s+\S+)', line)
                    if rate_match:
                        extracted_data["Rate"] = rate_match.group(1)

                # Scenario 5: Find exact word "Demand KW" and save the second number after that as "Const"
                if "Demand KW" in line:
                    demand_kw_match = re.findall(r'[\d,]+(\.\d+)?', line)
                    if len(demand_kw_match) >= 2:
                        if demand_kw_match and demand_kw_match[0]:
                            extracted_data["Demand KW"] = float(demand_kw_match[0].replace(',', ''))
                        else:
                            # Handle the case where demand_kw_match is empty or None
                            print("demand_kw_match is empty or None")
                            extracted_data["Demand KW"] = None  # or some default value
                        

                        # Find the second number after "Demand KW"
                        second_number_match = re.search(r'\d+(\.\d+)?\s+(\d+(\.\d+)?)', line)
                        if second_number_match:
                            extracted_data["Const"] = float(second_number_match.group(2).replace(',', ''))
                            constant_flag = True
                        
                            # Find the third number after "Demand KW" using various patterns
                        third_number_match = re.search(r'Demand KW.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'Demand KW.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'Demand KW.*?(\d+(\.\d+)?)', line)

                        if third_number_match:
                            # Check if group(5) exists in the match object before accessing it
                            if len(third_number_match.groups()) >= 5:
                                extracted_data["Usage"] = float(third_number_match.group(5).replace(',', ''))
                            elif len(third_number_match.groups()) >= 3:
                                extracted_data["Usage"] = float(third_number_match.group(3).replace(',', ''))
                            else:
                                # Handle the case where group(3) doesn't exist (set a default value or handle it as needed)
                                extracted_data["Usage"] = 0  # You can change this default value if needed
                
                regulatory_fee_match = re.search(r'Regulatory fee (State fee)', text, re.IGNORECASE)
                if regulatory_fee_match:
                                # If found, extract the number
                                extracted_data["Regulatory fee (State fee)"] = float(regulatory_fee_match.group(1).replace(',', ''))
                                #extracted_data["Regulatory fee (State fee)"] = float(regulatory_fee_match.group(1))
                                #print("mfjkdndnsjj")
                # Scenario 5: Find exact word "Demand KW" and save the second number after that as "Const"
                if "On-peak demand" in line:
                    demand_kw_match = re.findall(r'[\d,]+(\.\d+)?', line)
                    
                    if len(demand_kw_match) >= 2:
                        extracted_data["On-peak demand"] = float(demand_kw_match[0].replace(',', ''))

                        # Find the second number after "Demand KW"
                        second_number_match = re.search(r'\d+(\.\d+)?\s+(\d+(\.\d+)?)', line)
                        if second_number_match:
                            extracted_data["Const"] = float(second_number_match.group(2).replace(',', ''))
                            constant_flag = True

                            # Find the third number after "Demand KW" using various patterns
                        third_number_match = re.search(r'On-peak demand.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'On-peak demand.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'On-peak demand.*?(\d+(\.\d+)?)', line)

                        if third_number_match:
                            # Check if group(5) exists in the match object before accessing it
                            if len(third_number_match.groups()) >= 5:
                                extracted_data["On-peak demand2"] = float(third_number_match.group(5).replace(',', ''))
                            elif len(third_number_match.groups()) >= 3:
                                extracted_data["On-peak demand2"] = float(third_number_match.group(3).replace(',', ''))
                            else:
                                # Handle the case where group(3) doesn't exist (set a default value or handle it as needed)
                                extracted_data["On-peak demand2"] = 0  # You can change this default value if needed
                
                                
                # Scenario 6: Find exact word "Service to" and save 3 words after that
                if "Service to" in line:
                    service_to_match = re.search(r'Service to\s+(\S+\s+\d{1,2},\s+\d{4})', line)
                    if service_to_match:
                        extracted_data["Service to"] = service_to_match.group(1)
                        service_to_date = extracted_data["Service to"]
                        # Parse the date string into a datetime object
                        service_date_obj = dt.datetime.strptime(service_to_date, '%b %d, %Y')
                        # Extract the month as an integer (1-12)
                        extracted_data["Service Month"] = service_date_obj.month
                        service_to_flag = True

                # Check for other keywords
                for keyword in keywords:
                    if keyword in line:
                        # Scenario 1: Find the first number next to the exact word
                        number_match = re.search(rf'{keyword}\s*([\d,]+(\.\d+)?)', line)
                        if number_match:
                            extracted_data[keyword] = float(number_match.group(1).replace(',', ''))

                        # Scenario 2: Find $ sign next to the exact word
                        dollar_match = re.search(rf'{keyword}\s*\$\s*([\d,]+(\.\d+)?)', line)
                        if dollar_match:
                            extracted_data[keyword] = float(dollar_match.group(1).replace(',', ''))

                        # Scenario 3: Find $ sign between parentheses next to the exact word
                        parentheses_match = re.search(rf'{keyword}\s*\((.*?)\)', line)
                        if parentheses_match:
                          content_inside_parentheses = parentheses_match.group(1)
                          dollar_inside_parentheses_match = re.search(r'\$\s*([\d,]+(\.\d+)?)', content_inside_parentheses)
                          if dollar_inside_parentheses_match:
                              extracted_data[keyword] = float(dollar_inside_parentheses_match.group(1).replace(',', ''))
                          else:
                              # No $ sign between parentheses, extract the number after parentheses
                              number_after_parentheses_match = re.search(r'\)\s*([\d,]+(\.\d+)?)', line)
                              if number_after_parentheses_match:
                                  extracted_data[keyword] = float(number_after_parentheses_match.group(1).replace(',', ''))
     # Inside the extract_data function, after checking for "Non-fuel energy charge:"
    found_non_fuel_energy_charge = False  # Initialize the flag
    found_on_peak = False  # Initialize the flag to track "On-peak"
    for line in text.split('\n'):
        if "Non-fuel energy charge:" in line:
            found_non_fuel_energy_charge = True
            continue  # Move to the next line to search for "On-peak"

        # Check for "On-peak" in the line, but only if we've found "Non-fuel energy charge:"
        if found_non_fuel_energy_charge and "On-peak" in line:
            on_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
            
            if on_peak_match:
                # Save the matched number under the key "Non-fuel energy charge: on-peak"
                extracted_data["Non-fuel energy charge: on-peak"] = float(on_peak_match.group(1).replace(',', ''))
                found_on_peak = True  # Set the flag to indicate "On-peak" was found
            continue  # Move to the next line to search for "Off-peak"

        # Check for "Off-peak" in the line, but only if "On-peak" was found in the previous line
        if found_on_peak and "Off-peak" in line:
            off_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
            if off_peak_match:
                # Save the matched number under the key "Non-fuel energy charge: off-peak"
                extracted_data["Non-fuel energy charge: off-peak"] = float(off_peak_match.group(1).replace(',', ''))
                found_non_fuel_energy_charge = False  # Reset the flag
                found_on_peak = False  # Reset the "On-peak" flag

    
    
    
        
        Demand_charge = False
        found_demand_charge = False
        found_on_peak1 = False
        for line in text.split('\n'):
            if "Demand charge:" in line:
                found_demand_charge = True
                # Try to extract the number from the line
                demand_charge_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
                if demand_charge_match:
                    # Extract the value of "Demand charge"
                    extracted_data["Demand charge:"] = float(demand_charge_match.group(1).replace(',', ''))
                    
                    break  # Exit the loop after finding and extracting "Demand charge" value
                continue  # Move to the next line to search for "On-Peak" after "Demand charge" is found

            if found_demand_charge and "On-Peak" in line:
                on_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
                if on_peak_match:
                    extracted_data["Demand charge-On-peak"] = float(on_peak_match.group(1).replace(',', ''))
                    found_on_peak1 = True
                    break  # Exit the loop after finding "On-peak" since you only want to extract it once
        
        
        Demand_charge_On_peak = extracted_data.get("Demand charge-On-peak", 0)
        #print(Demand_charge_On_peak)
        
        for line in text.split('\n'):
            non_fuel_charge_match = re.search(r'Non-fuel energy charge:\s*\n\s*\$([\d,]+(\.\d+)?)', line)

            # Check if a match is found
            if non_fuel_charge_match:
                # Extract the matched number and convert it to a float
                non_fuel_charge_value = float(non_fuel_charge_match.group(1).replace(',', ''))
                extracted_data["Non-fuel energy charge:"] = non_fuel_charge_value

        lines = text.split('\n')
        for i in range(len(lines)):
            # Check if the line contains the exact phrase "Non-fuel energy charge:"
            if "Non-fuel energy charge:" in lines[i]:
                # Extract the number from the next line
                next_line = lines[i+1].strip()
                non_fuel_charge_match = re.search(r'\$([\d,]+(\.\d+)?)', next_line)

                # Check if a match is found
                if non_fuel_charge_match:
                    # Extract the matched number and convert it to a float
                    non_fuel_charge_value = float(non_fuel_charge_match.group(1).replace(',', ''))
                    extracted_data["Non-fuel energy charge:"] = non_fuel_charge_value
                    
    # Inside the extract_data function, after checking for "Fuel charge:"
        found_Fuel_charge= False  # Initialize the flag
        found_on_peak = False  # Initialize the flag to track "On-peak"
        for line in text.split('\n'):
            fuel_charge_match = re.search(r'Fuel charge: \$([\d,]+(\.\d+)?)', text)

            if fuel_charge_match:
        # Extract the matched number and convert it to a float
                fuel_charge_value = float(fuel_charge_match.group(1).replace(',', ''))
                
                # Assign the extracted value to the dictionary
                extracted_data["Fuel charge:"] = fuel_charge_value
            if "Fuel charge:" in line:
                found_Fuel_charge = True
                continue  # Move to the next line to search for "On-peak"

            # Check for "On-peak" in the line, but only if we've found "Non-fuel energy charge:"
            if found_Fuel_charge and "On-peak" in line:
                
                on_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
                if on_peak_match:
                    # Save the matched number under the key "Non-fuel energy charge: on-peak"
                    extracted_data["Fuel charge-On-peak"] = float(on_peak_match.group(1).replace(',', ''))
                    found_on_peak = True  # Set the flag to indicate "On-peak" was found
                continue  # Move to the next line to search for "Off-peak"
            
            
           

            


            # Check for "Off-peak" in the line, but only if "On-peak" was found in the previous line
            if found_on_peak and "Off-peak" in line:
                off_peak_match = re.search(r'\s*([+-]?\d+(\.\d+)?)', line)
                if off_peak_match:
                    # Save the matched number under the key "Non-fuel energy charge: off-peak"
                    extracted_data["Fuel charge-Off-peak"] = float(off_peak_match.group(1).replace(',', ''))
                    found_Fuel_charge = False  # Reset the flag
                    found_on_peak = False  # Reset the "On-peak" flag

            

            pattern = r'FPL SolarTogether credit\s*([−\-\d,.]+)'

            




            # Search for the pattern in the text
            match = re.search(pattern, text)
            #print(match)
            if match:
                # If a match is found, extract the number including its sign
                #print("llllllllllllllllllllllllllllllllllllllllllllllllllll")
                # Replacing non-standard minus sign with standard minus sign
                FPL_SolarTogether_credit1 = (match.group(1))
                standard_string = FPL_SolarTogether_credit1.replace('−', '-')

                # Removing comma
                FPL_SolarTogether_credit = FPL_SolarTogether_credit1.replace(',', '')
                numeric_part = re.sub(r'[^\d.-]', '', FPL_SolarTogether_credit)

                # Converting to float
                FPL_SolarTogether_credit = float(numeric_part)
                # If the original string had a negative sign, multiply the float value by -1
                if '−' in FPL_SolarTogether_credit1:
                    FPL_SolarTogether_credit *= -1
                extracted_data["FPL SolarTogether credit"] = FPL_SolarTogether_credit
                #print(FPL_SolarTogether_credit)
          # Inside the extract_data function, after extracting all other values
    rate = str(extracted_data.get("Rate", "")).strip().upper()
    valid_rates = ["GSLDT-1 GENERAL", "GSDT-1 GENERAL", "GSLD-1 GENERAL", "HLFT-2 HIGH","HLFT-2 HIGH LOAD FACTOR DEMAND TIME OF USE","HLFT-1 HIGH","OL-1 OUTDOOR"]

    print(rate)
    if rate in valid_rates:
        
    #if "GSLDT-1 GENERAL" in rate or "GSDT-1 GENERAL" in rate or "GSLDT-1 GENERAL" in rate or "GSLD-1 GENERAL" in rate or "HLFT-2 HIGH LOAD FACTOR DEMAND TIME OF USE":
        print("rate hlfgfgf")
        non_fuel_off_peak = extracted_data.get("Non-fuel energy charge: off-peak", 0)
        fuel_off_peak = extracted_data.get("Fuel charge-Off-peak", 0)
        off_peak_kwh_used = extracted_data.get("Off-peak kWh used", 0)
        non_fuel_on_peak = extracted_data.get("Non-fuel energy charge: on-peak", 0)
        fuel_on_peak = extracted_data.get("Fuel charge-On-peak", 0)
        #on_peak_kwh_used = extracted_data.get("On-peak kWh used", 0)
        demand_charge = extracted_data.get("Demand charge:", 0)
        Demand_charge_On_peak = extracted_data.get("Demand charge-On-peak", 0)
        
        if demand_charge==0:
           demand_charge=Demand_charge_On_peak
           
        on_peak_demand2 = extracted_data.get("On-peak demand2", 0)
        
        
        on_peak_demand1 = extracted_data.get("On-peak demand", 0)
        
        Power_monitoring_premium_plus = extracted_data.get("Power monitoring-premium plus", 0)
        maximum_demand = extracted_data.get("Maximum demand", 0)
        maximum = extracted_data.get("Maximum", 0)
        franchise_charge = extracted_data.get("Franchise charge", 0)
        utility_tax =extracted_data.get("Utility tax", 0)
        florida_sales_tax =extracted_data.get("Florida sales tax", 0)
        gross_receipts_tax =extracted_data.get("Gross rec. tax/Regulatory fee", 0)
        gross_receipts_tax1 =extracted_data.get("Gross receipts tax", 0)
        county_sales_tax =extracted_data.get("County sales tax", 0)
        base_charge =extracted_data.get("Base charge:", 0)
        Reg_fee =extracted_data.get("Regulatory fee", 0)
        Discretionary_sales =extracted_data.get("Discretionary sales surtax", 0)
        Service_Charge =extracted_data.get("Service Charge", 0)
        franchise_fee = extracted_data.get("Franchise fee", 0)
        total_comsuption_kwh=extracted_data.get("kWh Used", 0)
        Late_payment_charge =extracted_data.get( "Late payment charge", 0)
        FPL_SolarTogether_charge =extracted_data.get( "FPL SolarTogether charge", 0)
        FPL_SolarTogether_credit =extracted_data.get( "FPL SolarTogether credit", 0)
        #print(FPL_SolarTogether_credit)
        #print("inam solar inam solar")
        Total_Comsuption_kwh= total_comsuption_kwh
        if 'Total_Comsuption_kwh' in locals() and Total_Comsuption_kwh is not None and Total_Comsuption_kwh != "":
            print("jjjjjjjjjjjjjjjjjjj")
            if off_peak_kwh_used != 0 and 'Off-peak kWh used' is not None:
               
               

   

            
                print("lolololpoloplop")
                print(off_peak_kwh_used)
                if Total_Comsuption_kwh!=0 :
                    
                    on_peak_kwh_used=total_comsuption_kwh-off_peak_kwh_used
                    print("popopopopo")
                    print(on_peak_kwh_used)
                    Total_comsuption_kwh =off_peak_kwh_used + on_peak_kwh_used
                    if gross_receipts_tax == Reg_fee :
                        Reg_fee=0
                    
                    if on_peak_demand2==0:
                        on_peak_demand=on_peak_demand1
                        
                    else: 
                        on_peak_demand=on_peak_demand2
                    
                    Total_Services_Tax=Reg_fee +base_charge + Discretionary_sales+Service_Charge+franchise_fee+county_sales_tax+gross_receipts_tax1+franchise_charge+ utility_tax  + florida_sales_tax + gross_receipts_tax+Late_payment_charge +Power_monitoring_premium_plus+FPL_SolarTogether_charge+FPL_SolarTogether_credit
                    
                    
                    Total_comsuption_kwh = on_peak_kwh_used + off_peak_kwh_used
                    
                    
                    Energy_Charge= non_fuel_on_peak*on_peak_kwh_used+non_fuel_off_peak*off_peak_kwh_used
                    Energy_Charge_On_peak=non_fuel_on_peak*on_peak_kwh_used
                    Energy_Charge_Off_peak= non_fuel_off_peak*off_peak_kwh_used
                    Fuel_Charge= fuel_off_peak* off_peak_kwh_used + fuel_on_peak* on_peak_kwh_used
                    
                    Fuel_Charge_on_peak=fuel_on_peak* on_peak_kwh_used
                    Fuel_Charge_off_peak=fuel_off_peak* off_peak_kwh_used
                    Total_Energy_Charge=Energy_Charge+Fuel_Charge
                    Total_dolar_khw=Total_Energy_Charge/Total_comsuption_kwh
                    On_Peak_demand_Charge=demand_charge * on_peak_demand
                    Maximum_demand_Charge=maximum_demand * maximum
                    Total_Demand_Charge= On_Peak_demand_Charge + Maximum_demand_Charge
                    Total_Electric_cost=Total_Energy_Charge+ Total_Demand_Charge
                    
                    Total_Charge=Total_Electric_cost+ Total_Services_Tax
                
                    Energy_Rate= (Total_Energy_Charge)/(Total_comsuption_kwh)
                    Demand_Rate=(Total_Demand_Charge)/(maximum_demand +maximum)
                    extracted_data["Power monitoring-premium plus"] = Power_monitoring_premium_plus
                    extracted_data["Energy Charge"] = Energy_Charge
                    extracted_data["Energy Charge On peak"] = Energy_Charge_On_peak
                    extracted_data["Energy Charge Off peak"] = Energy_Charge_Off_peak
                    extracted_data["Fuel Charge"] = Fuel_Charge
                    extracted_data["Fuel Charge on peak $"] = Fuel_Charge_on_peak
                    extracted_data["Fuel Charge off peak $"] = Fuel_Charge_off_peak
                    extracted_data["Total Energy Charge"] = Total_Energy_Charge
                    extracted_data["Total Electric cost"] = Total_Electric_cost
                    extracted_data["Total Services and Tax"] = Total_Services_Tax
                    extracted_data["Total Charge"] = Total_Charge
                    extracted_data["Total $/kwh cost"] = Total_dolar_khw
                    extracted_data["On Peak Demand Charge"] = On_Peak_demand_Charge
                    
                    extracted_data["Maximum Demand Charge"] = Maximum_demand_Charge
                    extracted_data["Total Demand Charge TOU ($)"] = Total_Demand_Charge
                    extracted_data["Total Comsuption KWH"] = Total_comsuption_kwh
                    extracted_data["Demand charge:"] = Demand_charge_On_peak
                    if Demand_charge_On_peak==0:
                        extracted_data["Demand charge:"] =demand_charge
                    extracted_data["Total Demand"] =on_peak_demand
                    extracted_data["Energy Rate"] = Energy_Rate
                    extracted_data["Demand Rate"] = Demand_Rate
                    extracted_data["On-Peak kWh used"] = on_peak_kwh_used
                    if on_peak_demand2==0:
                        on_peak_demand2=on_peak_demand
                        extracted_data["On-peak demand2"] = on_peak_demand2
                    
                
         
        on_peak_kwh_used=extracted_data.get('On-peak kWh used')
        print("lllllllllllll")
        print(on_peak_kwh_used)
        if off_peak_kwh_used ==0:
            
            non_fuel = extracted_data.get("Non-fuel energy charge:", 0)
            fuel = extracted_data.get("Fuel charge:", 0)
            
            if non_fuel==0:
                non_fuel=(non_fuel_off_peak+non_fuel_on_peak)/2
                
            if fuel==0:
                fuel=(fuel_on_peak+fuel_off_peak)/2 
            
            #fuel = extracted_data.get("Fuel:", 0)
            print(non_fuel)
            #demand_kw = extracted_data.get("Demand KW", 0)
            demand = extracted_data.get("Demand charge:", 0)
            kwh_used = extracted_data.get("kWh Used", 0)
            
                                       
            usage1 = extracted_data.get("Demand KW", 0)
            usage2 = extracted_data.get("Usage", 0)
            print(usage1)
            print(usage2)
            print("emeoooooooooooooooooooooooooooooooooooooooo")
            if usage1 is None:
               usage1 = 0
            if usage2 is None:
               usage2 = 0

            usage=max(usage1,usage2)
            if usage is None or usage == 0:
               usage = extracted_data.get("Usage", 0)
            
            Total_Comsuption_kwh=kwh_used
            print(Total_Comsuption_kwh)
            base_charge =extracted_data.get("Base charge:", 0)
            Customer_charge =extracted_data.get("Customer charge:", 0)
            print(base_charge)
            if base_charge==0:
                base_charge=Customer_charge
            Reg_fee =extracted_data.get("Regulatory fee", 0)
            if Reg_fee == 0:
                Reg_fee =extracted_data.get("Regulatoiy fee (State fee)", 0)
            
            Gross_reciep =extracted_data.get("Gross receipts tax", 0)
            Gross_rec =extracted_data.get("Gross rec. tax/Regulatory fee", 0)
            if Gross_reciep:
                if Reg_fee==Gross_reciep :
                    Reg_fee=0
                    extracted_data["Gross receipts tax"] = 0
            if Reg_fee:
                if Gross_rec == Reg_fee:
                    Reg_fee=0
            Energy_Charge=kwh_used*non_fuel
            Fuel_Charge= kwh_used * fuel
            print(Fuel_Charge)
            print(kwh_used)
            
            
            
            
            utility_tax =extracted_data.get("Utility tax", 0)
            franchise_fee = extracted_data.get("Franchise fee", 0)
            franchise_charge = extracted_data.get("Franchise charge", 0)
            
            
            #Total_Comsuption_kwh=extracted_data.get("kWh Used", 0)
            florida_sales_tax =extracted_data.get("Florida sales tax", 0)
            Discretionary_sales =extracted_data.get("Discretionary sales surtax", 0)
            county_sales_tax =extracted_data.get("County sales tax", 0)
            Contract_demand =extracted_data.get("Contract demand", 0)
            Late_payment_charge =extracted_data.get( "Late payment charge", 0)
            FPL_SolarTogether_charge =extracted_data.get( "FPL SolarTogether charge", 0)
            FPL_SolarTogether_credit =extracted_data.get( "FPL SolarTogether credit", 0)
            print(usage)
            print(demand)
            print("regfeee")
            print(Reg_fee)
            if on_peak_kwh_used !=0:
               kwh_used=Total_Comsuption_kwh        
            
            Total_Energy_Charge=Energy_Charge+Fuel_Charge
            if Contract_demand !=0:
                
               Total_Demand_Charge= usage * demand + Contract_demand * demand
            else: 
                Total_Demand_Charge=usage * demand 
            Total_Electric_cost=Total_Energy_Charge+ Total_Demand_Charge
            Total_Services_Tax=Gross_rec + Gross_reciep + utility_tax + franchise_fee + franchise_charge+base_charge + Reg_fee+florida_sales_tax+Discretionary_sales+county_sales_tax+Late_payment_charge+FPL_SolarTogether_credit+FPL_SolarTogether_charge+Power_monitoring_premium_plus
            Total_Charge=Total_Electric_cost+ Total_Services_Tax
            Energy_Rate= (Total_Energy_Charge)/(kwh_used)
            Demand_Rate=(Total_Demand_Charge)/(usage)
            Total_dolar_khw=Total_Energy_Charge/Total_Comsuption_kwh
            if usage==0:
                usage=1
            extracted_data["Energy Charge"] = Energy_Charge
            extracted_data["Fuel Charge"] = Fuel_Charge
            extracted_data["Total Energy Charge"] = Total_Energy_Charge
            extracted_data["Total Electric cost"] = Total_Electric_cost
            extracted_data["Total Services and Tax"] = Total_Services_Tax
            extracted_data["Total Charge"] = Total_Charge
            extracted_data["Total Comsuption KWH"] = Total_Comsuption_kwh
            extracted_data["Total Energy Charge"] = Total_Energy_Charge
            extracted_data["Total Demand Charge - Non TOU ($)"] = Total_Demand_Charge
            if Contract_demand!=0:
                Contract_demand=Contract_demand-1
            extracted_data["Total Demand"] = usage + Contract_demand
            extracted_data["Energy Rate"] = Energy_Rate
            extracted_data["Demand Rate"] = Demand_Rate
            extracted_data["Total $/kwh cost"] = Total_dolar_khw
            extracted_data["Total $/kwh cost"] = Total_dolar_khw
    # Inside the extract_data function, after extracting all other values
    else:
      print("dorostedoroste")
      extracted_data.get("Rate", "") == "GSD-1 GENERAL"
      
      non_fuel = extracted_data.get("Non-fuel:", 0)
      fuel = extracted_data.get("Fuel:", 0)
      #demand_kw = extracted_data.get("Demand KW", 0)
      demand = extracted_data.get("Demand:", 0)
      kwh_used = extracted_data.get("kWh Used", 0)
      usage = extracted_data.get("Usage", 0)
      base_charge =extracted_data.get("Base charge:", 0)
      Gross_rec =extracted_data.get("Gross rec. tax/Regulatory fee", 0)
      Gross_reciep =extracted_data.get("Gross receipts tax", 0)
      utility_tax =extracted_data.get("Utility tax", 0)
      franchise_fee = extracted_data.get("Franchise fee", 0)
      franchise_charge = extracted_data.get("Franchise charge", 0)
      Reg_fee =extracted_data.get("Regulatory fee", 0)
      Customer_charge =extracted_data.get("Customer charge:", 0)
      #Total_Comsuption_kwh=extracted_data.get("kWh Used", 0)
      florida_sales_tax =extracted_data.get("Florida sales tax", 0)
      Discretionary_sales =extracted_data.get("Discretionary sales surtax", 0)
      county_sales_tax =extracted_data.get("County sales tax", 0)
      Contract_demand =extracted_data.get("Contract demand", 0)
      Late_payment_charge =extracted_data.get( "Late payment charge", 0)
      
            
      Total_Comsuption_kwh=kwh_used
      
      if base_charge==0:
         base_charge=Customer_charge

      if Gross_reciep:
        if Reg_fee==Gross_reciep :
            Reg_fee=0
            extracted_data["Gross receipts tax"] = 0
      if Reg_fee:
        if Gross_rec == Reg_fee:
            Reg_fee=0
      Energy_Charge=kwh_used*non_fuel
      Fuel_Charge= kwh_used * fuel
      Total_Energy_Charge=Energy_Charge+Fuel_Charge
      Total_Demand_Charge= usage * demand + Contract_demand * demand
      Total_Electric_cost=Total_Energy_Charge+ Total_Demand_Charge
      Total_Services_Tax=Gross_rec + Gross_reciep + utility_tax + franchise_fee + franchise_charge+base_charge + Reg_fee+florida_sales_tax+Discretionary_sales+county_sales_tax+Late_payment_charge
      Total_Charge=Total_Electric_cost+ Total_Services_Tax
      Energy_Rate= (Total_Energy_Charge)/(kwh_used)
      Total_dolar_khw=Total_Energy_Charge/Total_Comsuption_kwh
      if usage==0:
         usage=1
      Demand_Rate=(Total_Demand_Charge)/(usage)
      extracted_data["Energy Charge"] = Energy_Charge
      extracted_data["Fuel Charge"] = Fuel_Charge
      extracted_data["Total Energy Charge"] = Total_Energy_Charge
      extracted_data["Total Electric cost"] = Total_Electric_cost
      extracted_data["Total Services and Tax"] = Total_Services_Tax
      extracted_data["Total Charge"] = Total_Charge
      extracted_data["Total Comsuption KWH"] = Total_Comsuption_kwh
      extracted_data["Total Energy Charge"] = Total_Energy_Charge
      extracted_data["Total Demand Charge - Non TOU ($)"] = Total_Demand_Charge
      if Contract_demand!=0:
          Contract_demand=Contract_demand-1
      extracted_data["Total Demand"] = usage + Contract_demand
      extracted_data["Energy Rate"] = Energy_Rate
      extracted_data["Demand Rate"] = Demand_Rate
      extracted_data["Total $/kwh cost"] = Total_dolar_khw

      

    # If "Service to" was not found in the current PDF, mark it as NaN
    if not service_to_flag:
        extracted_data["Service to"] = float('nan')

    return extracted_data


keywords=["Rate", "Service to","Service days", "Total Comsuption KWH", "Energy Charge", "Fuel:", "Fuel Charge","Fuel Charge on peak $","Fuel Charge off peak $", "Non-fuel:", "Energy Charge On peak","Energy Charge Off peak","Total Energy Charge", "Total $/kwh cost",
                 "Usage", "Total Demand Charge - Non TOU ($)","Total Demand Charge TOU ($)","Contract demand", "Total Electric cost", "Base charge:", "Gross rec. tax/Regulatory fee", "Franchise charge", "Franchise fee", "Utility tax",
                 "Florida sales tax", "Discretionary sales surtax", "Taxes and charges", "Gross receipts tax", "Regulatory fee","Regulatory fee (State fee)", "County sales tax", "Service Charge", "On-Peak kWh used", 
                 "Off-peak kWh used", "On-peak demand","FPL SolarTogether charge","FPL SolarTogether credit", "Maximum demand","Demand KW","kWh Used","Demand charge:","Maximum","Non-fuel energy charge: on-peak","Late payment charge",
                 "Non-fuel energy charge: off-peak","Regulatoiy fee (State fee)", "Fuel charge-On-peak", "Fuel charge-Off-peak", "Total Charge", "Energy Rate", "Demand Rate", "Demand:","Customer charge:","On-peak demand2","Power monitoring-premium plus"]
# Define keywords to search for

# Create a directory to store text files
if not os.path.exists("text_files"):
    os.makedirs("text_files")

# Create a list to store data for each account
data_by_account = []

# Create a list to store the PDF file names
pdf_file_names = []

# Create a set to keep track of processed PDF file names
processed_pdf_files = set()

# Create a dictionary to accumulate the values across all accounts
cumulative_data = {
    "Taxes and charges_A": 0,
    "Total charges": 0,
    "Total Comsuption KWH": 0,
    "Total Energy Charge": 0,
    "Total Demand Charge": 0,
}

def extract_and_consolidate_data(uploaded_files, num_accounts, coefficients):
    # List to hold data for all accounts
    data_by_account = []

    for account_number in range(num_accounts):
        data_for_account = []
        months_present = []  # List to store months present for this account

        # Loop through each uploaded file
        for uploaded_file in uploaded_files:
            pdf_path = uploaded_file.name  # Use the uploaded file's name as the path
            # Read the file data
            extracted_data = extract_data(uploaded_file)
            data_for_account.append(extracted_data)
            print(f"Extracted data from {pdf_path} for Account {account_number + 1}")
            # Extract the "Service Month" from the extracted data
            service_month = extracted_data.get("Service Month")
            if service_month:
                months_present.append(service_month)

        data_by_account.append(data_for_account)
            
            

    # Create an Excel file with separate sheets for each account
    excel_filename = 'all_accounts_data.xlsx'
    with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as excel_writer:
        data_by_account_transposed = []  # List to store transposed data for each account
           
        missing_months_by_account = []
        for i, account_data in enumerate(data_by_account):
            account_sheet_name = f'Account_{i + 1}'
            df_account = pd.DataFrame(account_data)
            # Transpose the DataFrame
            #check_values = calculate_check_values(start_month)
            missing_months = []  # Initialize the variable
            # Transpose the DataFrame
            df_account_transposed = df_account.transpose()
            print(df_account_transposed)
            # Check if "Service Month" is in the index (rows)
            if "Service Month" in df_account_transposed.index:
                # Extract the "Service Month" row
                service_month_row = df_account_transposed.loc["Service Month"]

                # Define a set of all months from 1 to 12
                all_months = set(range(1, 13))

                # Extract the actual service months as integers
                service_months = {int(month) for month in service_month_row.values if str(month).isdigit()}
                
                # Find the missing months
                missing_months_account = sorted(list(all_months - service_months))
                missing_months_account1=missing_months_account
                # Now you have the missing months for this account
                print(f"Missing months for {account_sheet_name}: {missing_months_account}")
                if not missing_months_account1:
                # Get the values in the "Service Month" row
                    service_month_values = df_account_transposed.loc["Service Month"]

                    # Rename the columns with the corresponding values
                    df_account_transposed.columns = service_month_values
                miss_NH = set()         # Initialize a set for missing months without both previous and next months               
                index_of_average=0
                
                column_values_month = []  # To store the result
                for month in missing_months_account1:
                    
                    previous_month = month - 1
                    next_month = month + 1
                    column_values_month=[]
                    if previous_month== 0:
                        previous_month=12
                    if next_month==13:
                        next_month=1
                    
                    
                    if previous_month not in missing_months_account1 and next_month not in missing_months_account1:
                        miss_NH.add(month)
                        
                        column_name_p = df_account_transposed.columns[df_account_transposed.loc["Service Month"] == previous_month].item()
                        column_values_p = df_account_transposed[column_name_p]
                        
                        # Convert next_month to an integer
                        next_month = int(next_month)
                        if next_month==13: 
                            next_month=1
                        # Check if next_month exists in the "Service Month" row
                        if next_month in df_account_transposed.loc["Service Month"].values:
                            # Get the corresponding column name
                            column_name_n = df_account_transposed.columns[df_account_transposed.loc["Service Month"] == next_month].item()
                            column_values_n = df_account_transposed[column_name_n]
                            #print(column_values_n)
                        else:
                            print(f"Next month {next_month} not found in 'Service Month'.")
                        
                        
                        for n, p in zip(column_values_n, column_values_p):
                            try:
                                # Try to convert n and p to float and perform the operation
                                n_float = float(n)
                                p_float = float(p)
                                result = (n_float + p_float) / 2
                                column_values_month.append(result)
                            except ValueError:
                                # Handle non-numeric values
                                column_values_month.append(None)  # You can use None or any other value as needed
                                
                        
                        number_of_members = len(missing_months_account1)
                        index_of_average=12-number_of_members 
                        #column_name = str(index_of_average)
                        column11=index_of_average+month
                        #print(column11)
                        if column11== 12:
                            column11=112
                        df_account_transposed[column11] = column_values_month
                        df_account_transposed.at["Late payment charge", column11] = 0
                        df_account_transposed.at["Service Month",column11] = month
                        
                        
                        missing_months_account1.remove(month)
                # Find the exact column name for "12"
                column_name = None
                for col in df_account_transposed.columns:
                    if col == 18:
                        column_name = col
                        break
                # Get the values in the "Service Month" row
                service_month_values = df_account_transposed.loc["Service Month"]

                # Rename the columns with the corresponding values
                df_account_transposed.columns = service_month_values
                            
                if column_name is not None:
                    # Change the value of "Service Month" in the found column to 12
                    df_account_transposed.at["Service Month", column_name] = 12
                    #print("jujujujuju")
                    #print(df_account_transposed)
                print("Missing months without both previous and next months:", miss_NH)
                missing_months_account = list(set(missing_months_account1) - miss_NH)
                print(missing_months_account1)  
                #print("iiii") 
                if not missing_months_account1 :
                    print("empty")
                else:          
                    # Flatten the list of coefficients
                    flat_coefficients = [coeff[0] for coeff in coefficients]
                    # Find the maximum coefficient value and its month
                    max_coefficient = max(flat_coefficients)
                    max_coefficient_month = flat_coefficients.index(max_coefficient) + 1    
                    service_month_row = df_account_transposed.loc["Service Month"]
                    # Check if 1 is missing
                    if max_coefficient_month not in missing_months_account and missing_months_account:
                        # Find the row where "Service Month" is 1
                        index_of_1 = service_month_row[service_month_row == max_coefficient_month].index[0]
                        # Extract the corresponding column
                        based_anchorm_co = df_account_transposed[index_of_1]
                        
                        
                        value1 = df_account_transposed.loc["Demand:",  index_of_1]
                        value2 = df_account_transposed.loc["Fuel:",  index_of_1]
                        value3 = df_account_transposed.loc["Fuel Charge on peak $",  index_of_1]
                        value4 = df_account_transposed.loc["Fuel Charge off peak $",  index_of_1]
                        value5 = df_account_transposed.loc["Non-fuel:",  index_of_1]
                        value6 = df_account_transposed.loc["Base charge:",  index_of_1]
                        value7 = df_account_transposed.loc["Gross rec. tax/Regulatory fee",  index_of_1]
                        value8 = df_account_transposed.loc[ "Franchise charge",  index_of_1]
                        value9 = df_account_transposed.loc["Franchise fee",  index_of_1]
                        value10 = df_account_transposed.loc["Utility tax",  index_of_1]
                        value11 = df_account_transposed.loc["Florida sales tax",  index_of_1]
                        value12 = df_account_transposed.loc["Discretionary sales surtax",  index_of_1]
                        value13 = df_account_transposed.loc["Gross receipts tax",  index_of_1]
                        value14 = df_account_transposed.loc["Regulatory fee",  index_of_1]
                        value15 = df_account_transposed.loc["County sales tax",  index_of_1]
                        value16 = df_account_transposed.loc["Service Charge",  index_of_1]
                        value17 = df_account_transposed.loc["Maximum",  index_of_1]
                        value18 = df_account_transposed.loc["Demand charge:",  index_of_1]
                        value19 = df_account_transposed.loc["Non-fuel energy charge: on-peak",  index_of_1]
                        value20 = df_account_transposed.loc["Non-fuel energy charge: off-peak", index_of_1]
                        value21 = df_account_transposed.loc["Fuel charge-On-peak",  index_of_1]
                        value22 = df_account_transposed.loc["Fuel charge-Off-peak",  index_of_1]
                        value23 = df_account_transposed.loc["Customer charge:", index_of_1]
                        value24 = df_account_transposed.loc["On-peak demand2",  index_of_1]
                        value25 = df_account_transposed.loc["Service days",  index_of_1]
                        
                    
                        # Modify the DataFrame for missing months
                        
                        for missing_month in missing_months_account:
                            
                        
                            def multiply_by_constant(value):
                                    try:
                                        missing_month1 = missing_month - 1
                                        coefficient = float(coefficients[missing_month1][0])
                                        print(f"Coefficient for missing month {missing_month1}: {coefficient}")
                                        
                                        return float(value) * coefficient
                                        
    
                                    except (ValueError, TypeError):
                                        return value

                                # Apply the function to each element in the DataFrame
                            based_anchorm_co12 = based_anchorm_co.apply(multiply_by_constant)
                            service_month_values = df_account_transposed.loc["Service Month"]
                            df_account_transposed.columns = service_month_values
                            df_account_transposed[missing_month] = based_anchorm_co12   
                            df_account_transposed.at["Late payment charge", missing_month] = 0
                            df_account_transposed.at["Demand:", missing_month] = value1
                            df_account_transposed.at["Fuel:", missing_month] = value2
                            df_account_transposed.at["Fuel Charge on peak $", missing_month] = value3
                            df_account_transposed.at["Fuel Charge off peak $", missing_month] = value4
                            df_account_transposed.at["Non-fuel:", missing_month] = value5
                            df_account_transposed.at["Base charge:", missing_month] = value6
                            df_account_transposed.at["Gross rec. tax/Regulatory fee", missing_month] = value7
                            df_account_transposed.at["Franchise charge", missing_month] = value8
                            df_account_transposed.at["Franchise fee", missing_month] = value9
                            df_account_transposed.at["Utility tax", missing_month] = value10
                            df_account_transposed.at["Florida sales tax", missing_month] = value11
                            df_account_transposed.at["Discretionary sales surtax", missing_month] = value12
                            df_account_transposed.at["Gross receipts tax", missing_month] = value13
                            df_account_transposed.at["Regulatory fee", missing_month] = value14
                            df_account_transposed.at["County sales tax", missing_month] = value15
                            df_account_transposed.at["Service Charge", missing_month] = value16
                            df_account_transposed.at["Maximum", missing_month] = value17
                            df_account_transposed.at["Demand charge:", missing_month] = value18
                            df_account_transposed.at["Non-fuel energy charge: on-peak", missing_month] = value19
                            df_account_transposed.at["Non-fuel energy charge: off-peak", missing_month] = value20
                            df_account_transposed.at["Fuel charge-On-peak", missing_month] = value21
                            df_account_transposed.at["Fuel charge-Off-peak", missing_month] = value22
                            df_account_transposed.at["Customer charge:", missing_month] = value23
                            df_account_transposed.at["On-peak demand2", missing_month] = value24
                            df_account_transposed.at["Service days", missing_month] = value25
                            df_account_transposed.at['Service Month', missing_month] = missing_month                   
                            
                            

                                            
                        
                    elif max_coefficient_month  in missing_months_account and missing_months_account:
                            alpha = next(i for i in range(1, 13) if i not in missing_months_account)
                            print(f"Selected alpha: {alpha}") 
                            # Find the row where "Service Month" is equal to alpha
                            index_of_alpha = service_month_row[service_month_row == alpha].index[0]
                            value1 = df_account_transposed.loc["Demand:", index_of_alpha]
                            value2 = df_account_transposed.loc["Fuel:", index_of_alpha]
                            value3 = df_account_transposed.loc["Fuel Charge on peak $", index_of_alpha]
                            value4 = df_account_transposed.loc["Fuel Charge off peak $",index_of_alpha]
                            value5 = df_account_transposed.loc["Non-fuel:", index_of_alpha]
                            value6 = df_account_transposed.loc["Base charge:",index_of_alpha]
                            value7 = df_account_transposed.loc["Gross rec. tax/Regulatory fee", index_of_alpha]
                            value8 = df_account_transposed.loc[ "Franchise charge", index_of_alpha]
                            value9 = df_account_transposed.loc["Franchise fee", index_of_alpha]
                            value10 = df_account_transposed.loc["Utility tax", index_of_alpha]
                            value11 = df_account_transposed.loc["Florida sales tax",index_of_alpha]
                            value12 = df_account_transposed.loc["Discretionary sales surtax", index_of_alpha]
                            value13 = df_account_transposed.loc["Gross receipts tax", index_of_alpha]
                            value14 = df_account_transposed.loc["Regulatory fee", index_of_alpha]
                            value15 = df_account_transposed.loc["County sales tax", index_of_alpha]
                            value16 = df_account_transposed.loc["Service Charge", index_of_alpha]
                            value17 = df_account_transposed.loc["Maximum", index_of_alpha]
                            value18 = df_account_transposed.loc["Demand charge:", index_of_alpha]
                            value19 = df_account_transposed.loc["Non-fuel energy charge: on-peak", index_of_alpha]
                            value20 = df_account_transposed.loc["Non-fuel energy charge: off-peak", index_of_alpha]
                            value21 = df_account_transposed.loc["Fuel charge-On-peak", index_of_alpha]
                            value22 = df_account_transposed.loc["Fuel charge-Off-peak", index_of_alpha]
                            value23 = df_account_transposed.loc["Customer charge:",index_of_alpha]
                            value24 = df_account_transposed.loc["On-peak demand2", index_of_alpha] 
                            value25 = df_account_transposed.loc["Service days",  index_of_alpha]    
                            
                            
                            # Modify the DataFrame for missing months
                            # Extract the corresponding column
                            based_anchorm_co1 = df_account_transposed[index_of_alpha]
                            
                            # Get the values in the "Service Month" row
                            service_month_values = df_account_transposed.loc["Service Month"]

                            # Rename the columns with the corresponding values
                            df_account_transposed.columns = service_month_values
                            
                            
                            for missing_month in missing_months_account:
                                
                                
                                    def multiply_by_constant(value):
                                        try:
                                            missing_month1 = missing_month - 1
                                            coefficient = float(coefficients[missing_month1][0])
                                            print(f"Coefficient for missing month {missing_month1}: {coefficient}")
                                            
                                            return float(value) * coefficient
                                        

                                        
                                        except (ValueError, TypeError):
                                            return value
                                
                                    

                                        # Apply the function to each element in the DataFrame
                                    based_anchorm_co122 = based_anchorm_co1.apply(multiply_by_constant)
                                    df_account_transposed[missing_month] = based_anchorm_co122
                                    df_account_transposed.at["Late payment charge", missing_month] = 0 
                                    df_account_transposed.at["Demand:",  missing_month] = value1
                                    df_account_transposed.at["Fuel:",  missing_month] = value2
                                    df_account_transposed.at["Fuel Charge on peak $",  missing_month] = value3
                                    df_account_transposed.at["Fuel Charge off peak $",  missing_month] = value4
                                    df_account_transposed.at["Non-fuel:",  missing_month] = value5
                                    df_account_transposed.at["Base charge:",  missing_month] = value6
                                    df_account_transposed.at["Gross rec. tax/Regulatory fee",  missing_month] = value7
                                    df_account_transposed.at["Franchise charge",  missing_month] = value8
                                    df_account_transposed.at["Franchise fee",  missing_month] = value9
                                    df_account_transposed.at["Utility tax", missing_month] = value10
                                    df_account_transposed.at["Florida sales tax",  missing_month] = value11
                                    df_account_transposed.at["Discretionary sales surtax",  missing_month] = value12
                                    df_account_transposed.at["Gross receipts tax",  missing_month] = value13
                                    df_account_transposed.at["Regulatory fee",  missing_month] = value14
                                    df_account_transposed.at["County sales tax",  missing_month] = value15
                                    df_account_transposed.at["Service Charge",  missing_month] = value16
                                    df_account_transposed.at["Maximum",  missing_month] = value17
                                    df_account_transposed.at["Demand charge:", missing_month] = value18
                                    df_account_transposed.at["Non-fuel energy charge: on-peak",  missing_month] = value19
                                    df_account_transposed.at["Non-fuel energy charge: off-peak",  missing_month] = value20
                                    df_account_transposed.at["Fuel charge-On-peak",  missing_month] = value21
                                    df_account_transposed.at["Fuel charge-Off-peak",  missing_month] = value22
                                    df_account_transposed.at["Customer charge:",  missing_month] = value23
                                    df_account_transposed.at["On-peak demand2", missing_month] = value24
                                    df_account_transposed.at["Service days", missing_month] = value25
                                    value_at_position12 = df_account_transposed.at["Service to", missing_month]
                                    column_names = df_account_transposed.columns                          

                                    # Assuming missing_month is a column
                                    column_name = str(missing_month)
                                                                
                                    # Assuming df_account_transposed is your DataFrame
                                    # Update the "Service to" row with the new month
                                
                            month_map = {
                                            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr',
                                            5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug',
                                            9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
                                        }             
                                                
                            # Extract the column names as integers
                            column_names = df_account_transposed.columns

                            # Convert column names to month names using the provided 'month_map'
                            month_names = [month_map[column_name] for column_name in column_names]

                        # Ensure "Service to" row is a datetime series
                            df_account_transposed.loc["Service to", :] = pd.to_datetime(df_account_transposed.loc["Service to", :], errors='coerce')

                            # Extract the day and year from the existing datetime values
                            day_year_str = df_account_transposed.loc["Service to", :].apply(lambda x: x.strftime('%d, %Y') if pd.notna(x) else '')

                            # Replace NaN values (non-datetime) with the new month_names
                            df_account_transposed.loc["Service to", :] = day_year_str + ' ' + month_names

                            # Ensure "Service to" row is a datetime series again
                            df_account_transposed.loc["Service to", :] = pd.to_datetime(df_account_transposed.loc["Service to", :], errors='coerce')


                            print(f"Based on 'Service Month' = {alpha}: {based_anchorm_co1}")
               
            
            
            
            #row_label = "Service to"
            #column_names = df_account_transposed.columns

            # Now, let's find the corresponding values in each column for the specified row
            #values_for_row = df_account_transposed.loc[row_label]               
            #value_at_position = df_account_transposed.at["Service to", 3]
            print(missing_months_account)
            #print("injaaa")
            if missing_months_account:
                                
                # Extract the values of "Service Month"
                service_month_values = df_account_transposed.loc["Service Month"].astype(int)

                # Sort the columns based on "Service Month" values
                sorted_columns = service_month_values.sort_values().index

                # Sort the DataFrame based on sorted columns
                df_account_transposed = df_account_transposed[sorted_columns]
                
                 
            else:
                
            # Extract the 'Service to' values from the transposed DataFrame
                service_to_values = pd.to_datetime(df_account_transposed.loc["Service to"], format='%b %d, %Y')

            # Sort the columns based on the "Service to" dates
                sorted_columns = service_to_values.sort_values().index
                            
                df_account_transposed = df_account_transposed[sorted_columns]
            
            # Calculate the sum of each row excluding "Service to" and "Rate"
            excluded_columns = ["Service to", "Rate"]
            sums = df_account_transposed.drop(excluded_columns, axis=0).sum(axis=1)

            # Add a new "Sum" column with the sums
            df_account_transposed["Sum"] = sums
            
            # Append the sorted and modified DataFrame to the list
            row_to_update_demand_rate = df_account_transposed.loc["Demand:"]
                        
            data_by_account_transposed.append(df_account_transposed)
            
            consolidate_transpose_df = pd.concat(data_by_account_transposed).groupby(level=0).sum()
            #print(consolidate_transpose_df)         
            # Filter the DataFrame to obtain the row with the label "Demand charge:"
            demand_charge_row = consolidate_transpose_df.loc["Demand charge:"]
            #print(demand_charge_row)
            print(consolidate_transpose_df)
            # Extract the values from the desired column (replace 'column_name' with the actual column name)
            value = demand_charge_row[11]  # Replace 'column_name' with the actual column name
            column_values = value
            consolidate_transpose_df.at["Demand charge:", 'Sum'] = column_values
            consolidate_transpose_df.loc["Demand:"] =row_to_update_demand_rate
            
            # Filter the DataFrame to obtain the row with the label "Demand charge:"
            Maximum_row = consolidate_transpose_df.loc["Maximum"]
            value = Maximum_row[11]  # Replace 'column_name' with the actual column name
            column_values = value
            # Extract the values from the desired column (replace 'column_name' with the actual column name)
            #column_values = demand_charge_row[11].tolist()
            consolidate_transpose_df.at["Maximum", 'Sum'] = column_values
            consolidate_transpose_df.loc["Service days"] = consolidate_transpose_df.loc["Service days"]/num_accounts
            
            # Calculate the new value
            new_value = (consolidate_transpose_df.at["Energy Charge", "Sum"] / consolidate_transpose_df.at["Total Comsuption KWH", "Sum"])

            # Update the value in the DataFrame
            consolidate_transpose_df.at["Non-fuel:", "Sum"] = new_value

            # Calculate the new value
            new_value = (consolidate_transpose_df.at["Fuel Charge", "Sum"] / consolidate_transpose_df.at["Total Comsuption KWH", "Sum"])

            # Update the value in the DataFrame
            consolidate_transpose_df.at["Fuel:", "Sum"] = new_value
            
            consolidate_transpose_df.loc["Total Demand Charge"]=consolidate_transpose_df.loc[ "Total Demand Charge - Non TOU ($)"]+consolidate_transpose_df.loc["Total Demand Charge TOU ($)"]
            
            #if not consolidate_transpose_df.loc["Usage"].empty and (consolidate_transpose_df.loc["Usage"] != 0).any():
            if not consolidate_transpose_df.loc["Usage"].empty and not (consolidate_transpose_df.loc["Usage"] == 0).any():
                valueeee=consolidate_transpose_df.loc["Usage"]
                consolidate_transpose_df.loc["Demand:"] = consolidate_transpose_df.loc["Total Demand Charge - Non TOU ($)"]/consolidate_transpose_df.loc["Usage"]
            # Calculate the values for the new row "Demand $/kwh"
            demand_per_kwh = consolidate_transpose_df.loc["Total Demand Charge"] / consolidate_transpose_df.loc["Total Demand"]
            # Add the new row to the DataFrame
            consolidate_transpose_df.loc["Demand $/kwh"] = demand_per_kwh
            consolidate_transpose_df.loc["Non-TOU Consumption KWH"]=consolidate_transpose_df.loc["Total Comsuption KWH"]-consolidate_transpose_df.loc["On-Peak kWh used"]-consolidate_transpose_df.loc["Off-peak kWh used"]
            print(consolidate_transpose_df.loc["Non-TOU Consumption KWH"])
            print(consolidate_transpose_df.loc["On-Peak kWh used"])
            print(consolidate_transpose_df.loc["Total Comsuption KWH"])
            print("lklkinjoooojahah")

            consolidate_transpose_df.loc["Energy Charge Non-TOU ($)"]=consolidate_transpose_df.loc["Energy Charge"]-consolidate_transpose_df.loc["Energy Charge On peak"]-consolidate_transpose_df.loc["Energy Charge Off peak"]
            if not consolidate_transpose_df.loc["Non-TOU Consumption KWH"].empty and (consolidate_transpose_df.loc["Non-TOU Consumption KWH"] != 0).any():
               consolidate_transpose_df.loc["Energy $/kwh on Non-TOU"]=consolidate_transpose_df.loc["Energy Charge Non-TOU ($)"]/consolidate_transpose_df.loc["Non-TOU Consumption KWH"]

            if not consolidate_transpose_df.loc["Fuel Charge"].empty and (consolidate_transpose_df.loc["Fuel Charge"]!=0).any():
               consolidate_transpose_df.loc["Fuel Charge Non-TOU $"]=consolidate_transpose_df.loc["Fuel Charge"]-consolidate_transpose_df.loc["Fuel Charge on peak $"]-consolidate_transpose_df.loc["Fuel Charge off peak $"]
               
            print("innnnnnnn")
            print(consolidate_transpose_df)

            if not consolidate_transpose_df.loc["Non-TOU Consumption KWH"].empty and (consolidate_transpose_df.loc["Non-TOU Consumption KWH"] != 0).any():
               consolidate_transpose_df.loc["Fuel $/KWH Non-TOU"]=consolidate_transpose_df.loc["Fuel Charge Non-TOU $"]/consolidate_transpose_df.loc["Non-TOU Consumption KWH"]

            
            consolidate_transpose_df.loc["Fuel:"] = consolidate_transpose_df.loc["Fuel Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc["Non-fuel:"] = consolidate_transpose_df.loc["Energy Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc[ "Total $/kwh cost"] = consolidate_transpose_df.loc["Total Energy Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc[ "Energy Rate"] = consolidate_transpose_df.loc["Total Energy Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc[ "Demand Rate"] = consolidate_transpose_df.loc["Total Demand Charge"] / consolidate_transpose_df.loc["Total Demand"]
            
            
               
            if consolidate_transpose_df.at["On-Peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Energy Charge On peak", "Sum"] / consolidate_transpose_df.at["On-Peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Non-fuel energy charge: on-peak", "Sum"] = new_value 
            
            if consolidate_transpose_df.at["Off-peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Energy Charge Off peak", "Sum"] / consolidate_transpose_df.at["Off-peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Non-fuel energy charge: off-peak", "Sum"] = new_value 
            
            
            if consolidate_transpose_df.at["On-Peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Fuel Charge on peak $", "Sum"] / consolidate_transpose_df.at["On-Peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Fuel charge-On-peak", "Sum"] = new_value 
            
            if consolidate_transpose_df.at["Off-peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Fuel Charge off peak $", "Sum"] / consolidate_transpose_df.at["Off-peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Fuel charge-Off-peak", "Sum"] = new_value 
                
            

            # Remove rows with "Rate" and "Service to" if present
            if "Rate" in consolidate_transpose_df.index:
                consolidate_transpose_df.drop("Rate", inplace=True)
            if "Service to" in consolidate_transpose_df.index:
                consolidate_transpose_df.drop("Service to", inplace=True)
                
            
           # Create a list to store the sum values for each quarter
            Total_amount_per_Qtr = []

            # Define the number of columns in each quarter
            columns_per_quarter = 3

            # Find the index of the "Total Charge" row
            total_charge_index = consolidate_transpose_df.index.get_loc("Total Charge")

            # Iterate through the DataFrame by selecting columns for each quarter
            
            for i in range(0, len(consolidate_transpose_df.columns), columns_per_quarter):
                quarter = consolidate_transpose_df.columns[i:i + columns_per_quarter]
                quarter_sum = consolidate_transpose_df.iloc[total_charge_index, i:i + columns_per_quarter].sum()
                Total_amount_per_Qtr.append(quarter_sum)
                
                
            # Ensure 'Quarter_sum' has 12 values, inserting zeros where needed
            while len(Total_amount_per_Qtr) < 13:
                Total_amount_per_Qtr.append(0)
            # Now, 'quarter_sum_df' contains the 'Quarter_sum' values
            
            # Convert 'Quarter_sum' to a DataFrame with the same columns as the original DataFrame
            quarter_sum_df = pd.DataFrame([Total_amount_per_Qtr], columns=consolidate_transpose_df.columns, index=["Total amount per Qtr2"])

            # Concatenate the 'quarter_sum_df' with the original DataFrame to add it as a new row
            consolidate_transpose_df = pd.concat([consolidate_transpose_df, quarter_sum_df])
            row_to_update = consolidate_transpose_df.loc["Total amount per Qtr2"]
            max_column = row_to_update.idxmax()
            consolidate_transpose_df.at["Total amount per Qtr2",max_column]=0
            
            # Create a list to hold the values for the new row
            new_row_values = []
            new22=consolidate_transpose_df.loc["Total amount per Qtr2"]
            # Convert the Series to a list
            new22_list = new22.tolist()

            # Find indices of non-zero elements
            non_zero_indices = [i for i, value in enumerate(new22_list) if value != 0]

            # Get corresponding column names from the DataFrame
            column_names = consolidate_transpose_df.columns[non_zero_indices]
            first_four_values = [new22_list[i] for i in non_zero_indices[:4]]

            
            # Convert the Series to a list
            new22_list = new22.tolist()

            # Initialize the values for "pw"
            pw_values = [0] * 13

            # Find indices of non-zero elements
            non_zero_indices = [i for i, value in enumerate(new22_list) if value != 0]

            # Assign the non-zero values to the appropriate positions in "pw"
            for i, index in enumerate([2, 5, 8, 11]):
                if i < len(non_zero_indices):
                    pw_values[index] = new22_list[non_zero_indices[i]]

            # Create a new row "pw" in the DataFrame with the values
            consolidate_transpose_df.loc["Total amount per Qtr"] = pw_values
             
            
            # Create a mapping from month numbers to month names
            month_mapping = {
                1: 'Jan',
                2: 'Feb',
                3: 'Mar',
                4: 'Apr',
                5: 'May',
                6: 'Jun',
                7: 'Jul',
                8: 'Aug',
                9: 'Sep',
                10: 'Oct',
                11: 'Nov',
                12: 'Dec',
                'Sum': 'Sum'
            }
            
            
            # Rename the columns in the DataFrame using the mapping
            consolidate_transpose_df.columns = [month_mapping[col] for col in consolidate_transpose_df.columns]
            

            
            desired_order=["Service days",  "On-Peak kWh used", "Off-peak kWh used","Non-TOU Consumption KWH","kWh Used"," ","Energy Charge On peak","Energy Charge Off peak","Energy Charge Non-TOU ($)", "Energy Charge"," ",  
                "Non-fuel energy charge: on-peak","Non-fuel energy charge: off-peak","Energy $/kwh on Non-TOU","Non-fuel:"," ","Fuel Charge on peak $","Fuel Charge off peak $","Fuel Charge Non-TOU $", "Fuel Charge"," ","Fuel charge-On-peak", "Fuel charge-Off-peak","Fuel $/KWH Non-TOU","Fuel:"," ","Total Energy Charge", "Total $/kwh cost"," ",
                  "Usage","Contract demand", "On-peak demand2","Maximum demand","Total Demand"," ","Demand:","Demand charge:","Maximum"," ","Total Demand Charge - Non TOU ($)","Total Demand Charge TOU ($)", "Total Demand Charge", "Demand $/kwh"," ","Total Electric cost", " ","Base charge:","Service Charge", "Late payment charge"," ",
                  "Gross rec. tax/Regulatory fee", "Franchise charge", "Franchise fee", "Utility tax",
                 "Florida sales tax", "Discretionary sales surtax","FPL SolarTogether charge","FPL SolarTogether credit" , "Gross receipts tax", "Regulatory fee", "County sales tax", 
                 "Total Services and Tax"," ","Power monitoring-premium plus", "",

                   "Total Charge","Total amount per Qtr", " " ,"Energy Rate", "Demand Rate"]
            consolidate_transpose_df = consolidate_transpose_df.reindex(desired_order, axis=0)
            off_peak_kwh_used = extracted_data.get("Off-peak kWh used", 0)
            if off_peak_kwh_used !=0:
                 rename_dict = {
                 "Demand charge:" : "On-peak Demand $/kwh"}
            # Define a dictionary to map old row names to new row names
            rename_dict = {
                "Usage": "Total Demand kw - Non TOU", "Fuel :":"Fuel Charge $/kwh" , "Energy Charge" :"Total Energy Charge ($)",
                "Energy Charge On peak":"Energy Charge On peak ($)","Energy Charge Off peak": "Energy Charge Off peak ($)","Non-TOU Consumption KWH":"Non-TOU Consumption kwh",
                "Non-fuel:": "Average Energy $/kWh", "On-peak demand2":"On-peak demand kw","Contract demand" :"Contract demand kw",
                "On-Peak kWh used": "Consumption On Peak kwh","Fuel Charge" : "Total Fuel Charge ($)","Fuel Charge on peak $":"Fuel Charge on peak ($)","Fuel Charge off peak $":"Fuel Charge off peak ($)",
                "Off-peak kWh used": "Consumption off-Peak kwh","Total Energy Charge":"Total Energy & Fuel Charge ($)",
                "Demand:" : "Demand_$/kwh- Non TOU" , "Fuel:": "Average Fuel $/kWh",
                "Non-fuel energy charge: on-peak" : " Energy $/kwh on peak","Fuel Charge Non-TOU $":"Fuel Charge Non-TOU ($)",
                "Non-fuel energy charge: off-peak" :  " Energy $/kwh off peak","Total Electric cost":"Total Electric cost ($)",
                "Fuel charge-On-peak" : "Fuel Charge $/kwh on peak","Total Demand Charge":"Total Demand Charge ($)",
                 "Fuel charge-Off-peak" : "Fuel Charge $/kwh off peak","Total Demand": "Total Demand kw",
                 "Maximum demand" : "Maximum Demand kw" , "On-peak demand" : "On-peak Demand",
                 "Maximum" : "Maximum Demand $/kwh", "Base charge:" :"Base charge($)","Service Charge" : "Service Charge ($)",
                 "Late payment charge": "Late payment charge ($)",
                 "kWh Used" : "Total Comsuption kwh", "Energy Rate" : "Average $/kwh cost (Exc fees)", "Total Charge":"Total Charge ($)",
                 "Total Services and Tax": "Total Services and Tax ($)",
                 "Gross rec. tax/Regulatory fee": "Gross rec. tax/Regulatory fee ($)", "Franchise charge":"Franchise charge ($)", 
                 "Franchise fee": "Franchise fee ($)", "Utility tax":"Utility tax ($)","Florida sales tax":"Florida sales tax ($)", 
                 "Discretionary sales surtax":"Discretionary sales surtax ($)", "Taxes and charges":"Taxes and charges ($)", "Gross receipts tax": "Gross receipts tax ($)",
                 "Regulatory fee":"Regulatory fee ($)", "County sales tax":"County sales tax ($)","Total Services and Tax":"Total Services and Tax ($)",
                  "Total Charge":"Total Charge ($)","Total amount per Qtr":"Total amount per Qtr ($)"}
            
            # Use the rename method to change row names in the DataFrame
            consolidate_transpose_df = consolidate_transpose_df.rename(index=rename_dict)
             # Save the consolidated transpose data to the 'Consolidate Transpose' sheet
            #consolidate_transpose_df = consolidate_transpose_df.shift(periods=2, axis=1).shift(periods=3, axis=0)
            consolidate_transpose_df.to_excel(excel_writer, sheet_name='Consolidated')
            
            
            
            ###########AR
            ########################################################################
            implementation_cost_late=0
            total_annual_saving_late=consolidate_transpose_df.at["Late payment charge ($)", "Sum"] 
            simple_payback_late='Immediate'
            
            if total_annual_saving_late != 0:
                applicable="Applicable"
            else:
             applicable="Not Applicable"
              
            nan_value='#'
            ########################################################################
            # Find the value just before the maximum value/load factor
            
            A=consolidate_transpose_df.loc["Total Demand kw"].max()
            values_before_max = consolidate_transpose_df.loc["Total Demand kw"][consolidate_transpose_df.loc["Total Demand kw"] < A]
            # Find the maximum value from the values before the max
            value_before_max = values_before_max.max()
            B=consolidate_transpose_df.loc["Total Demand kw","Sum"]/12
            C=consolidate_transpose_df.loc[ "Demand Rate","Sum"]
            implementation_cost_load_factor=0
            AA= float(value_before_max)
            B = float(B)
            C = float(C)
            
            total_annual_saving_load_factor= (AA-B)*C
            simple_payback_load_factor='Immediate'
            ###########################################################################
            parameter=1.4
            Max_Demand_Expectation=parameter*consolidate_transpose_df.loc["Total Comsuption kwh","Sum"]/(12*consolidate_transpose_df.loc["Service days","Sum"])
            
            implementation_cost_Max_Demand_Expectation=0
            DDD=consolidate_transpose_df.loc["Total Demand kw"].min()
            if B> Max_Demand_Expectation:
               total_annual_saving_Max_Demand_Expectation=(B- Max_Demand_Expectation)*C 
            else :
                total_annual_saving_Max_Demand_Expectation=0
            simple_payback_Max_Demand_Expectation='Immediate'
            
            ###########################################################################
            
            Toatal_cost_TOU1=consolidate_transpose_df.loc["Energy Charge On peak ($)","Sum"]+consolidate_transpose_df.loc["Energy Charge Off peak ($)","Sum"]
            Toatal_cost_TOU2=consolidate_transpose_df.loc["Fuel Charge on peak ($)","Sum"]+consolidate_transpose_df.loc["Fuel Charge off peak ($)","Sum"]
            Total_demand_charge_TOU=consolidate_transpose_df.loc[ "Total Demand Charge TOU ($)","Sum"]
            
            Toatal_cost_TOU_GSDT= Toatal_cost_TOU1+Toatal_cost_TOU2+Total_demand_charge_TOU
            
            
            Total_comsuption_TOU=consolidate_transpose_df.loc["Consumption On Peak kwh","Sum"]+consolidate_transpose_df.loc["Consumption off-Peak kwh","Sum"]
            Total_rate_Fuel_Energy=consolidate_transpose_df.loc["Fuel $/KWH Non-TOU","Sum"]+consolidate_transpose_df.loc["Energy $/kwh on Non-TOU","Sum"]
            #print(consolidate_transpose_df)
            #print(Total_rate_Fuel_Energy)
            tolerance = 1e-6
            # Define the variable
            Total_rate_Fuel_Energy = np.nan  # Replace with your variable

            # Check if the variable is NaN
            if np.isnan(Total_rate_Fuel_Energy):
              Total_rate_Fuel_Energy=0.035
            Total_cost_comsuption_FE=Total_comsuption_TOU*Total_rate_Fuel_Energy
            Demand_rate_GSD=consolidate_transpose_df.loc["Demand_$/kwh- Non TOU","Sum"]
            #print(Demand_rate_GSD)
            if Demand_rate_GSD <tolerance:
               Demand_rate_GSD=11.25
            Total_cost_Demand_GSD = consolidate_transpose_df.loc["Maximum Demand kw","Sum"]*Demand_rate_GSD
            Total_cost_NON_TOU_GSD=Total_cost_comsuption_FE+Total_cost_Demand_GSD
            
            total_Annual_Saving_to_GSD=Toatal_cost_TOU_GSDT-Total_cost_NON_TOU_GSD
            implementation_cost_change_rate=0
            simple_payback_change_rate='Immediate'
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data1 = {
                "Category": ["Late payment fees were discovered upon examination of the electrical bills",
                             "Implementation Cost of Late",
                            "Total Annual Saving of Late",
                            "Simple Payback of Late",applicable
                            ],
                "$ Value": [nan_value,implementation_cost_late,
                        total_annual_saving_late,
                        simple_payback_late,nan_value
                        ]
            }
            
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data2 = {
                "Category": ["Load Factor: We can Reduce Max Demand and saving Cost",
                            "Implementation Cost of Load Factor",
                            "Total Annual Saving of Load Factor",
                            "Estimated Demand saving = (Max_Demand - Average_Demand)*Demand_Rate",applicable,
                            "Simple Payback of Load Factor"],
                "$ Value": [nan_value,implementation_cost_late,
                        total_annual_saving_load_factor,nan_value,
                        simple_payback_late,nan_value]
            }
            
            
            
            nan_value1="#"
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data3 = {
                "Category": ["Implementation Cost of Max Demand","total_annual_saving_Max_Demand_Expectation=(Average_Demand- 1.4*Max_Demand_Expectation)*Demand Rate",
                            "Total Annual Saving of Max Demand",
                            "Simple Payback of Max Demand",applicable],
                "$ Value": [
                        implementation_cost_Max_Demand_Expectation,nan_value1,
                        total_annual_saving_Max_Demand_Expectation,
                        simple_payback_Max_Demand_Expectation,nan_value]
            }
            
            
            
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data4 = {
                "Category": ["In order to switch to the GSD-1 rate","total_Annual_Saving_to_GSD=Toatal_cost_TOU_GSDT-Total_cost_NON_TOU_GSD",
                            "Implementation Cost of Change Rate",
                            "Total Annual Saving Change Rate",
                            "Simple Payback of Change Rate",applicable],
                "$ Value": [nan_value1,nan_value1,
                        implementation_cost_change_rate,
                       total_Annual_Saving_to_GSD,
                        simple_payback_change_rate,nan_value]
            }
            
            

            ad_model1=pd.DataFrame(data1)
            ad_model2=pd.DataFrame(data2)
            ad_model3=pd.DataFrame(data3)
            ad_model4=pd.DataFrame(data4)
            
            
            ad_model1.to_excel(excel_writer, sheet_name='Pay Electrical Bills On Time')
            ad_model2.to_excel(excel_writer, sheet_name='Load Factor')
            ad_model3.to_excel(excel_writer, sheet_name='Expectation of Max Demand')
            ad_model4.to_excel(excel_writer, sheet_name='Change Rate Structure to GSD')

             # Get the Excel writer's workbook and worksheet objects
            workbook = excel_writer.book
            worksheet = excel_writer.sheets['Consolidated']
            worksheet1 = excel_writer.sheets['Pay Electrical Bills On Time']
            worksheet2 = excel_writer.sheets['Load Factor']
            worksheet3 = excel_writer.sheets['Expectation of Max Demand']
            worksheet4 = excel_writer.sheets['Change Rate Structure to GSD']
         
            

             # Define the background color (e.g., green)
            green_fill = workbook.add_format({'bg_color': '00FF00'})

            # Define the row names to highlight
            rows_to_highlight = ["Average Energy $/kWh","Average Fuel $/kWh","Total Energy Charge ($)", "Total Demand Charge ($)",
                                 "Total Energy & Fuel Charge ($)","Total Fuel Charge ($)","Total $/kwh cost","Total Demand kw","Total Charge ($)","Total Services and Tax ($)", "Total Comsuption kwh","Total Electric cost ($)"]
            
            blue_fill = workbook.add_format({'bg_color': '#B4C6E7'})

                      
            consolidate_transpose_df = consolidate_transpose_df.fillna(0)


            # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_highlight:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, green_fill)  # Add +1 because row_num is zero-based
                        
                        
            rows_to_light = [ "Service days","Consumption On Peak kwh","Consumption off-Peak kwh","Non-TOU Consumption kwh","Energy Charge On peak ($)","Energy Charge Off peak ($)",
                             "Energy Charge Non-TOU ($)"," Energy $/kwh on peak"," Energy $/kwh off peak","Energy $/kwh on Non-TOU","Fuel Charge on peak ($)",
                             "Fuel Charge off peak ($)","Fuel Charge Non-TOU ($)","Fuel Charge $/kwh on peak","Fuel Charge $/kwh off peak","Fuel $/KWH Non-TOU","Total Demand kw - Non TOU",
                             "Contract demand kw","On-peak demand kw","Maximum Demand kw","Demand_$/kwh- Non TOU","On-peak Demand $/kwh","Maximum Demand $/kwh",
                             "Total Demand Charge - Non TOU ($)","Total Demand Charge TOU ($)","Demand $/kwh","Base charge($)","Service Charge ($)","Late payment charge ($)",
                             "Gross rec. tax/Regulatory fee ($)","Franchise charge ($)","Franchise fee ($)","Utility tax ($)","Florida sales tax ($)","Discretionary sales surtax ($)",
                             "Gross receipts tax ($)","Regulatory fee ($)","County sales tax ($)","Total amount per Qtr ($)","Average $/kwh cost (Exc fees)","Demand Rate"]
                
                        
        # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_light:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, blue_fill)  # Add +1 because row_num is zero-based
                        
           # Define the background color (e.g., light yellow)
            light_yellow_fill = workbook.add_format({'bg_color': '#FFFF00'})  
            rows_to_light1 = [ "Gross rec. tax/Regulatory fee ($)","Franchise charge ($)","Franchise fee ($)","Utility tax ($)","Florida sales tax ($)","Discretionary sales surtax ($)",
                             "Gross receipts tax ($)","Regulatory fee ($)","County sales tax ($)"]
           # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_light1:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, light_yellow_fill)  # Add +1 because row_num is zero-based
                        
           # Define the background color (e.g., red)
            red_fill = workbook.add_format({'bg_color': '#FF0000'})     
            rows_to_light2 = [ "Late payment charge ($)"]
            # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_light2:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, red_fill)  # Add +1 because row_num is zero-based  
                                          
           # Create a format for text justification (e.g., left alignment)
            justify_format = workbook.add_format({'align': 'left'})

            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet.set_column('A:Z', None, justify_format)

            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet.set_column('A:A', 50)  # Adjust '15' to your preferred width
            
            # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet.set_column('B:Z', None, center_align_format)
           
            
            
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet1.set_column('A:Z', None, justify_format)
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet1.set_column('B:B', 90)  # Adjust '15' to your preferred width
            worksheet1.set_column('C:C', 50)  # Adjust '15' to your preferred width
             # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet1.set_column('C:Z', None, center_align_format)
            
            

                        # Assuming total_annual_saving_late is your condition
            if total_annual_saving_late != 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet1.write(5, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet1.write(5, 1, 'Not Applicable', bold_format)  # Green bold text for column B    
                # Column C for row 4
                #worksheet1.write(5, 2, 'Applicable', workbook.add_format({'font_color': 'green'}))  # Green text for column C    
               
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet2.set_column('A:Z', None, justify_format)
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet2.set_column('B:B', 90)  # Adjust '15' to your preferred width
            worksheet2.set_column('C:C', 50)  # Adjust '15' to your preferred width
             # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet2.set_column('C:Z', None, center_align_format)
            
            
                        # Assuming total_annual_saving_late is your condition
            if total_annual_saving_load_factor > 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet2.write(6, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet2.write(6, 1, 'Not Applicable', bold_format)  # Green bold text for column B   
            
            
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet3.set_column('A:Z', None, justify_format)
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet3.set_column('B:B', 130)  # Adjust '15' to your preferred width
            worksheet3.set_column('C:C', 50)  # Adjust '15' to your preferred width
             # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet3.set_column('C:Z', None, center_align_format)
            
            
                         # Assuming total_annual_saving_late is your condition
            if total_annual_saving_Max_Demand_Expectation > 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet3.write(5, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet3.write(5, 1, 'Not Applicable', bold_format)  # Green bold text for column B   
            
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet4.set_column('A:Z', None, justify_format)
            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet4.set_column('B:B', 90)  # Adjust '15' to your preferred width
            worksheet4.set_column('C:C', 50)  # Adjust '15' to your preferred width
            # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
            worksheet4.set_column('C:Z', None, center_align_format)
            
             # Assuming total_annual_saving_late is your condition
            if total_Annual_Saving_to_GSD > 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet4.write(6, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet4.write(6, 1, 'Not Applicable', bold_format)  # Green bold text for column B  

            df_account.to_excel(excel_writer, sheet_name=account_sheet_name, index=False)
            
           
         
       # Save 'result_df' to the 'Consolidate' sheet
        #result_df.to_excel(excel_writer, sheet_name='Consolidate')
        for i, df_account_transposed in enumerate(data_by_account_transposed):
            account_sheet_name = f'Account_{i + 1}_Transposed'
            
            
            
            df_account_transposed.to_excel(excel_writer, sheet_name=account_sheet_name)
        

    print(f"Data saved to {excel_filename}")
    return excel_filename

    


# Streamlit App GUI
def app():
    st.title("FPL Bill PDF Extractor")
    
    # Input for number of accounts
    num_accounts = st.number_input("Enter the number of accounts:", min_value=1, step=1)
    
    # Define temperature coefficients for each month directly in the code
    #coefficients = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]  # Default coefficients for 12 months
    coefficients = [0.5904, 0.6416, 0.6672, 0.6672, 0.6928, 0.7183999999999999, 0.7952, 0.8464, 1.0, 0.7696, 0.744, 0.6672]

    # File uploader to upload multiple PDFs
    uploaded_files = st.file_uploader("Upload multiple FPL PDF files", type="pdf", accept_multiple_files=True)
    
    if st.button("Extract Data"):
        if uploaded_files and num_accounts > 0:
            # Call the function to process the PDFs and generate the Excel file
            excel_filename = extract_and_consolidate_data(uploaded_files, num_accounts, coefficients)
            
            # Provide download button for the Excel file
            with open(excel_filename, "rb") as file:
                btn = st.download_button(label="Download Excel File",
                                         data=file,
                                         file_name=excel_filename,
                                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Please upload at least one PDF file and specify the number of accounts.")

if __name__ == "__main__":
    app()
