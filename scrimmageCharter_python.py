# -*- coding: utf-8 -*-
"""
Created on Sat Dec 30 13:18:09 2023

@author: bennett stice
"""

import psycopg2

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
            Opponent VARCHAR(255)
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
            opponent = "Scrimmage"
            
            home_lineup= [int(x) for x in input("Enter the home team's lineup in number form with spaces between: ").split()]
            
            away_lineup= [int(x) for x in input("Enter the away team's lineup in number form with spaces between: ").split()]
            
            oppo_lineup = [away_lineup,home_lineup]

            h_fname,h_lname = input("Enter the first and last name of the home team's starting pitcher: ").split()
            a_fname,a_lname = input("Enter the first and last name of the away team's starting pitcher: ").split()
            fname = [h_fname,a_fname]
            lname = [h_lname,a_lname]

            inning = [1,1]
            outs = 0
            balls = 0
            strikes = 0
            pitch_count= [1,1]
            outs_accrued = [0,0]
            batter_in_inning = 1
            lineup_pos = [0,0]
            
            counter=0

            input("Hit Enter to start charting the game: ")

            go = True

            while go:
                print(f"Inning: {inning[counter%2]}   Pitcher: {fname[counter%2]} {lname[counter%2]}   Outs: {outs}   Count: {balls}-{strikes}  Batter #{oppo_lineup[counter%2][lineup_pos[counter%2]]}")
                
                user_input=input(">>> ")
                input_list=user_input.split()
                
                if len(input_list) not in (1,3):
                    print ("Invalid Entry. ENTER PITCH TYPE, VELOCITY, PITCH RESULT")
                
                elif user_input.lower() == "stop":
                    yes_no = input("Do you wish to stop charting? Y/N: ")
                    if yes_no.upper() == "Y":
                        go = False
                    else:
                        change = input("What do you wish to change? (Inning, Pitcher, Outs, Count, Batter): ")
                        if change.lower() == "inning":
                            inning[counter%2] = int(input("Enter the inning number: "))
                        elif change.lower() == "pitcher":
                            fname[counter%2], lname[counter%2] = input("Enter the first and last name of the pitcher: ").split()
                            pitch_count[counter%2] = 1
                            outs_accrued[counter%2]= 0
                        elif change.lower() == "outs":
                            prevOuts=outs
                            outs = int(input("Enter the amount of outs there are: "))
                            outs_accrued[counter%2]+=(outs-prevOuts)
                            if outs >= 3:
                                outs = 0
                                inning[counter%2] += 1
                                batter_in_inning = 1
                                counter+=1
                        elif change.lower() == "count":
                            balls, strikes = map(int, input("Enter the balls and strikes separated by a space: ").split())
                        elif change.lower() == "batter":
                            oppo_lineup[counter%2][lineup_pos[counter%2]] = int(input("Enter the number of the current batter: "))

                elif user_input.lower() == "undo":
                    prev_pitch_id=get_max_pitch_id(cursor)-1
                    prev_pitch_entry = "SELECT fname,lname,inning,outs,balls,strikes,outs_accrued,pitch_count,batter_of_inning,batter_number,pitch_id,ab_result,bip_result FROM pitch_log_T WHERE pitch_id=%s;"
                    cursor.execute(prev_pitch_entry,(prev_pitch_id,))
                    data = cursor.fetchone()
                    
                    if data[11]=="out" and outs==0:
                        counter-=1
                    
                    fname[counter%2]=data[0]
                    lname[counter%2]=data[1]
                    inning[counter%2]=data[2]
                    outs=data[3]
                    balls=data[4]
                    strikes=data[5]
                    outs_accrued[counter%2]=data[6]
                    pitch_count[counter%2]=data[7]
                    batter_in_inning=data[8]
                    batter_number=data[9]
                    pitch_id=data[10]
                    
                    if data[11]=="out":
                        outs_accrued[counter%2]-=1
                        if data[12]=="DP":
                            outs_accrued[counter%2]-=1
                    
                    if oppo_lineup[counter%2][lineup_pos[counter%2]]!=batter_number:
                        lineup_pos[counter%2]-=1
                        if lineup_pos[counter%2]==-1:
                            lineup_pos[counter%2]=len(oppo_lineup[counter%2])-1
                        
                    delete_statment = "DELETE FROM pitch_log_T WHERE pitch_id = %s;"
                    cursor.execute(delete_statment,(prev_pitch_id,))
                    print ("Previous Pitch Erased")

                elif len(input_list)!=3:
                    print ("Invalid Entry. ENTER PITCH TYPE, VELOCITY, PITCH RESULT")

                else:
                    pitch_type, velo, pitch_result = user_input.split()
                    
                    if pitch_type not in ("FB", "CH", "CU", "SL", "SP","CB"):
                        print("Invalid Pitch Type Entry.")
                        continue  # jump back to the start of the loop

                    try:
                        velo = float(velo)
                    except ValueError:
                        print("Invalid Velocity Entry")
                        continue  # jump back to the start of the loop

                    if pitch_result not in ("B", "SL", "SS", "SSC", "HBP", "D3SS", "BIP", "F"):
                        print("Invalid Pitch Result Entry")
                        continue  # jump back to the start of the loop
                    
                    BIP_result = input("What was the Ball-in-Play result: ") if pitch_result == "BIP" else "0"
                    
                    if BIP_result not in ("GO","FO","LO","1B","2B","3B","HR","E","SB","DP","0"):
                        print("Invalid Ball in Play Entry.")
                        continue #jump back to the start of the loop

                    if (pitch_result == "B" and balls == 3) or (
                            (pitch_result == "SL" or pitch_result == "SS" or pitch_result == "SSC") and strikes == 2) or (
                            pitch_result == "HBP") or (pitch_result == "D3SS") or (BIP_result != "0"):
                        if pitch_result == "B" or pitch_result == "HBP" or pitch_result == "D3SS" or BIP_result == "1B" or BIP_result == "2B" or BIP_result == "3B" or BIP_result == "HR" or BIP_result == "E":
                            AB_result = "safe"
                        else:
                            AB_result = "out"
                            outs_accrued[counter%2]+=1
                            if (BIP_result=="DP"):
                                outs_accrued[counter%2]+=1
                    else:
                        AB_result = "0"

                    pitch_log_inserter = """
                        INSERT INTO pitch_log_T (Pitch_ID, Fname, Lname, Inning, Outs, Balls, Strikes, Pitch_Type, Velocity, Pitch_Result, BIP_Result, Batter_Number, AB_result, Pitch_Count, Batter_Of_Inning, Outs_Accrued, Date, Opponent)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
                    """

                    cursor.execute(pitch_log_inserter, (pitch_id, fname[counter%2], lname[counter%2], inning[counter%2], outs, balls, strikes, pitch_type, velo, pitch_result, BIP_result, oppo_lineup[counter%2][lineup_pos[counter%2]], AB_result, pitch_count[counter%2], batter_in_inning, outs_accrued[counter%2], date, opponent))
                    connection.commit()

                    pitch_id += 1
                    pitch_count[counter%2] += 1

                    if pitch_result == "B":
                        balls += 1

                        if balls == 4:
                            balls = 0
                            strikes = 0
                            batter_in_inning += 1
                            lineup_pos[counter%2] += 1
                            if lineup_pos[counter%2] == len(oppo_lineup[counter%2]):
                                lineup_pos[counter%2] = 0

                    if pitch_result == "SL" or pitch_result == "SS" or pitch_result == "SSC":
                        strikes += 1

                        if strikes == 3:
                            balls = 0
                            strikes = 0
                            outs += 1
                            batter_in_inning += 1
                            lineup_pos[counter%2] += 1
                            if lineup_pos[counter%2] == len(oppo_lineup[counter%2]):
                                lineup_pos[counter%2] = 0

                            if outs == 3:
                                outs = 0
                                inning[counter%2] += 1
                                batter_in_inning = 1
                                counter+=1
                                

                    if pitch_result == "F":
                        if strikes < 2:
                            strikes += 1

                    if pitch_result == "HBP" or pitch_result == "D3SS":
                        balls = 0
                        strikes = 0
                        batter_in_inning += 1
                        lineup_pos[counter%2] += 1
                        if lineup_pos[counter%2] == len(oppo_lineup[counter%2]):
                            lineup_pos[counter%2] = 0

                    if BIP_result == "GO" or BIP_result == "FO" or BIP_result == "LO" or BIP_result == "SB" or BIP_result == "DP":
                        balls = 0
                        strikes = 0
                        outs += 1
                        batter_in_inning += 1
                        

                        if (BIP_result=="DP"):
                            outs += 1
                            
                        lineup_pos[counter%2] += 1
                        if lineup_pos[counter%2] == len(oppo_lineup[counter%2]):
                            lineup_pos[counter%2] = 0
                            
                        if outs >= 3:
                            outs = 0
                            inning[counter%2] += 1
                            batter_in_inning = 1
                            counter+=1
                            strikes=0
                            balls=0
                            
                            
                    if BIP_result=="1B" or BIP_result=="2B" or BIP_result=="3B" or BIP_result=="HR" or BIP_result=="E":
                        balls=0
                        strikes=0
                        batter_in_inning+=1
                        lineup_pos[counter%2]+=1
                        if lineup_pos[counter%2]==len(oppo_lineup[counter%2]):
                            lineup_pos[counter%2]=0
                        
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