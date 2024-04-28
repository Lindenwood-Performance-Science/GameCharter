# -*- coding: utf-8 -*-
"""
Created on Sat Apr 27 22:31:23 2024

@author: Bennett Stice
"""

import pandas as pd
import os
import psycopg2
from datetime import datetime

def create_pitch_log_table(cursor):
    pitch_log_table_creator = """
        CREATE TABLE IF NOT EXISTS pitch_log_T(
            Pitch_id SERIAL PRIMARY KEY,
            Fname VARCHAR(255),
            Lname VARCHAR(255),
            Inning INT,
            Outs INT,
            Balls INT,
            Strikes INT,
            Pitch_Type VARCHAR(255),
            Velocity INT,
            Pitch_Result VARCHAR(255),
            BIP_Result VARCHAR(255),
            Batter_Number INT,
            Outs_Accrued INT,
            AB_result VARCHAR(255),
            Pitch_Count INT,
            Batter_Of_Inning INT,
            Date VARCHAR(255),
            Opponent VARCHAR(255),
            Date_N Date
        );
    """
    cursor.execute(pitch_log_table_creator)

def get_max_pitch_id(cursor):
    pitch_id_setter = "SELECT max(Pitch_id) FROM pitch_log_T;"
    cursor.execute(pitch_id_setter)
    result = cursor.fetchone()
    return result[0] + 1 if result[0] is not None else 1

def main():
    try:
        connection = psycopg2.connect(
            dbname="ps1",
            user="pythoncon",
            password="password",
            host="18.217.248.114",
            port="5432"
        )

        with connection.cursor() as cursor:
            print(f"Opened database successfully: {connection.dsn}")

            create_pitch_log_table(cursor)
            pitch_id = get_max_pitch_id(cursor)
            connection.commit()

            date = input("Enter today's date (PUT IT IN MM-DD-YYYY form): ")
            date_n=datetime.strptime(date, '%m-%d-%Y')
            opponent = input("Enter today's opponent: ")

            oppo_lineup = [int(x) for x in input("Enter the opposing team's lineup in number form with spaces between: ").split()]

            fname, lname = input("Enter the first and last name of the pitcher: ").split()

            inning = 1
            outs = 0
            balls = 0
            strikes = 0
            pitch_count = 1
            outs_accrued = 0
            batter_in_inning = 1
            lineup_pos = 0
            
            columnslist=['Pitch_Type','Velocity','Pitch_Result','BIP_Result','STOP','Change','Update']
            #file=input("What is the excel file name: ")
            #df = pd.read_excel(file,usecols=columnslist)
            df = pd.read_excel('4-27-2024-WIU.xlsx',usecols=columnslist)
            
            for index, row in df.iterrows():
                
                    pitch_type=row['Pitch_Type']
                    velocity=row['Velocity']
                    pitch_result=row['Pitch_Result']
                    bip_result=row['BIP_Result']
                    stop_tag=row['STOP']
                    change=row['Change']
                    update=row['Update']
                    
                    
    
                    if stop_tag == "STOP":
                        if change == 'Pitcher':
                            fname,lname = update.split(" ")
                            pitch_count = 1
                            outs_accrued = 0
                        elif change == "Outs":
                            prevOuts=outs
                            outs=update
                            outs_accrued+=(outs-prevOuts)
                            if outs >= 3:
                                outs = 0
                                balls = 0
                                strikes = 0
                                inning += 1
                                batter_in_inning = 1
                        elif change == "Batter":
                            oppo_lineup[lineup_pos]=update
                        
                    else:
                        if pitch_type not in ("FF","CB","SL","CH", "CT","FT","SP"):
                            print("Invalid Pitch Type Entry.",index)
                            continue  # jump back to the start of the loop

                        try:
                            velocity = float(velocity)
                        except ValueError:
                            print("Invalid Velocity Entry")
                            continue  # jump back to the start of the loop

                        if pitch_result not in ("B", "SL", "SS", "SSC", "HBP", "D3SS", "BIP", "F"):
                            print("Invalid Pitch Result Entry",index)
                            continue  # jump back to the start of the loop
                        
                        
                        if bip_result not in ("GO","FO","LO","1B","2B","3B","HR","E","SB","DP","0"):
                            bip_result="0"

                        if (pitch_result == "B" and balls == 3) or (
                                (pitch_result == "SL" or pitch_result == "SS" or pitch_result == "SSC") and strikes == 2) or (
                                pitch_result == "HBP") or (pitch_result == "D3SS") or (bip_result != "0"):
                            if pitch_result == "B" or pitch_result == "HBP" or pitch_result == "D3SS" or bip_result == "1B" or bip_result == "2B" or bip_result == "3B" or bip_result == "HR" or bip_result == "E":
                                AB_result = "safe"
                            else:
                                AB_result = "out"
                                outs_accrued+=1
                                if (bip_result=="DP"):
                                    outs_accrued+=1
                        else:
                            AB_result = "0"

                        pitch_log_inserter = """
                            INSERT INTO pitch_log_T (Pitch_ID, Fname, Lname, Inning, Outs, Balls, Strikes, Pitch_Type, Velocity, Pitch_Result, BIP_Result, Batter_Number, AB_result, Pitch_Count, Batter_Of_Inning, Outs_Accrued, Date, Opponent,Date_N)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s);
                        """

                        cursor.execute(pitch_log_inserter, (pitch_id, fname, lname, inning, outs, balls, strikes, pitch_type, velocity, pitch_result, bip_result, oppo_lineup[lineup_pos], AB_result, pitch_count, batter_in_inning, outs_accrued, date, opponent,date_n))
                        #connection.commit()

                        pitch_id += 1
                        pitch_count += 1

                        if pitch_result == "B":
                            balls += 1

                            if balls == 4:
                                balls = 0
                                strikes = 0
                                batter_in_inning += 1
                                lineup_pos += 1
                                if lineup_pos == len(oppo_lineup):
                                    lineup_pos = 0

                        if pitch_result == "SL" or pitch_result == "SS" or pitch_result == "SSC":
                            strikes += 1

                            if strikes == 3:
                                balls = 0
                                strikes = 0
                                outs += 1
                                batter_in_inning += 1
                                lineup_pos += 1
                                if lineup_pos == len(oppo_lineup):
                                    lineup_pos = 0

                                if outs >= 3:
                                    outs = 0
                                    inning += 1
                                    batter_in_inning = 1
                                    strikes=0
                                    balls=0
                                    

                        if pitch_result == "F":
                            if strikes < 2:
                                strikes += 1

                        if pitch_result == "HBP" or pitch_result == "D3SS":
                            balls = 0
                            strikes = 0
                            batter_in_inning += 1
                            lineup_pos += 1
                            if lineup_pos == len(oppo_lineup):
                                lineup_pos = 0

                        if bip_result == "GO" or bip_result == "FO" or bip_result == "LO" or bip_result == "SB" or bip_result == "DP":
                            balls = 0
                            strikes = 0
                            outs += 1
                            batter_in_inning += 1
                            

                            if (bip_result=="DP"):
                                outs += 1
                                
                            lineup_pos += 1
                            if lineup_pos == len(oppo_lineup):
                                lineup_pos = 0
                                
                            if outs >= 3:
                                outs = 0
                                inning += 1
                                batter_in_inning = 1
                                
                                
                        if bip_result=="1B" or bip_result=="2B" or bip_result=="3B" or bip_result=="HR" or bip_result=="E":
                            balls=0
                            strikes=0
                            batter_in_inning+=1
                            lineup_pos+=1
                            if lineup_pos==len(oppo_lineup):
                                lineup_pos=0
                        
            connection.commit()
            
            
    except psycopg2.Error as e:
        # Handle database-related exceptions here
         print(f"Database error: {e}")

    except Exception as e:
          # Handle other exceptions here
          print(f"An unexpected error occurred: {e}")

    finally:
        # This block will be executed whether an exception occurs or not
        if connection:
            connection.close()
            print("Connection closed.")

if __name__ == "__main__":
    main()