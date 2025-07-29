
import os
import re
import datetime
#import subprocess
import time
import pyautogui
import pygetwindow as gw
import cx_Oracle
#import ctypes
#from ctypes import wintypes

#BlockInput = ctypes.windll.user32.BlockInput
#BlockInput.argtypes = [wintypes.BOOL]
#BlockInput.restype = wintypes.BOOL

#blocked = BlockInput(True)


#Definicje:
katalog=os.getcwd()
full_path_txt=str(katalog) + "\\" +'wynik.txt'
full_path_ezd=str(katalog) + '\\' +'laser.ezd'
#klawisz=''



#===================================================================================================================================
def zapisz_do_pliku(nazwa_pliku, dane):
    with open(nazwa_pliku, 'w') as plik:
        plik.write(dane)
        print('Zapisano dane do pliku: '+str(nazwa_pliku))



#===================================================================================================================================
def czy_plik_istnieje_i_aktualny(nazwa_pliku, czas_max_minuty=1):
    if os.path.exists(nazwa_pliku):
        mtime = os.path.getmtime(nazwa_pliku)
        mtime_dt = datetime.datetime.fromtimestamp(mtime)
        czas_teraz = datetime.datetime.now()
        roznica_czasu = czas_teraz - mtime_dt
        if roznica_czasu.total_seconds() / 60 <= czas_max_minuty:
            return True
    return False



#===================================================================================================================================
def multi_input():
    
    print("Chcesz wprowadzić kod QR z naklejki na detalu czy z przewodnika?")
    pyta1 = input()
    
    if pyta1.lower() in ["przewodnik"]:

        try:
            etykieta =''  
            lista = []
            licz = 1
            last = ''

            print("Wprowadź kod QR:")
            
            while True:
                #pyautogui.hotkey('alt', 'tab')
                wrd = input()

                if wrd.lower() in ["hala", "tak", "inna", "naklejka", "przewodnik"]:
                    os.system('cls' if os.name == 'nt' else 'clear')
                    break
                #print(wrd)
                #print("-$$$$$$-")
                if licz > 4 and len(wrd) < 1: break
                else:
                    if len(wrd) > 3:
                        lista.append(str(wrd))
                        licz = licz + 1
                    else:
                        if  wrd.isdigit() and len(wrd) > 3:
                            #print(wrd)
                            #print("-- !! --")
                            lista.append(str(wrd))
                            licz = licz + 1
                        else:
                            re1 = re.match("^PK ",last)
                            re2 = re.match("^PKA ",last)
                            if re1 or re2:
                                wymiar = re.findall("[0-9]{2,3}\*[0-9]{2,3}\*[0-9]{2,3}",last)
                                gatunek = re.findall("/(.*)/",last)
                                lista.append(wymiar[0].replace("*", "x"))
                                lista.append(str(gatunek[0]))
                                licz = licz + 1
                            else:
                                licz = licz + 1
                last = wrd
                #print(lista)
                #print("- ####### -")

        #What if ther user press the interruption shortcut? 
        except KeyboardInterrupt:
            print("program został zakończony przez użytkownika.")
            return

        if(len(lista)>0):
            #etykieta = '\n'.join(lista)
            etykieta = '\n'.join(lista)
            if len(etykieta) < 40:
                #raise Exception("cause of the problem")
                return ("99#9")
            else:
                    #PN = lista[len(lista)-1]
                    #PN = PN[4:]
                    #print(len(etykieta))
                    #print('Wprowadziłeś etykietę:')
                    #print(etykieta)
                print('\n')
                return   (etykieta)
        else:
            return ("99#9")
    if pyta1.lower() in ["naklejka"]:
        try:
            etykieta =''  
            lista = []
            licz = 1
            last = ''

            print("Wprowadź kod QR:")
            
            while True:
                #pyautogui.hotkey('alt', 'tab')
                wrd = input()

                if wrd.lower() in ["hala", "tak", "inna", "naklejka", "przewodnik"]:
                    os.system('cls' if os.name == 'nt' else 'clear')
                    break
                #print(wrd)
                #print("-$$$$$$-")
                #biore tekst, pierwszą częśc wrzucam do ORDER_NO, po znaku / do RELESEA_NO, po znaku / do SEQuENCE_NO
                if licz > 4 and len(wrd) < 1: break
                else:
                    if wrd.count("/") == 2:
                        if len(wrd) > 3:
                            lista.append(str(wrd))
                            #print(lista[0])
                            ORDER_no = ""
                            RELEASE_no = ""
                            SEQUENCE_no = ""
                            
                            czesci = wrd.split('/')

                            ORDER_no = czesci[0]
                            RELEASE_no = czesci[1]
                            SEQUENCE_no = czesci[2]
                            if len(czesci) != 3:
                                print("Wprowadzono niepoprawny tekst")
                                time.sleep(2)
                                return multi_input()
                            else:
                                #print( "PN: ", ORDER_no, RELEASE_no, SEQUENCE_no)
                                licz = licz + 1
                        else:
                            if  wrd.isdigit() and len(wrd) > 3:
                                ORDER_no = ""
                                RELEASE_no = ""
                                SEQUENCE_no = ""
                                
                                czesci = wrd.split('/')

                                ORDER_no = czesci[0]
                                RELEASE_no = czesci[1]
                                SEQUENCE_no = czesci[2]
                                
                                licz = licz + 1
                                #print( "PN: ", ORDER_no, RELEASE_no, SEQUENCE_no)
                            else:
                                licz = licz + 1
                    else: 
                        print("Wprowadzono niepoprawny tekst")
                        time.sleep(2)
                        return multi_input()        
                last = wrd

                


                # Połączenie z bazą danych
                #connection = cx_Oracle.connect("aszymkiewicz/ASZYMKIEWICZ@172.19.27.12:1521/szkol10")
                #dsn = cx_Oracle.makedsn(host='172.19.27.12', port=1521, sid='szkol10')
                #connection = cx_Oracle.connect(user='ASZYMKIEWICZ', password='ASZYMKIEWICZ', dsn=dsn)
                
                ip = '172.19.27.11'
                port = 1521
                service_name = 'PROD10'
                dsn = cx_Oracle.makedsn(ip, port, service_name=service_name)

                connection = cx_Oracle.connect('qlik', 'Q!propl!22@', dsn)

                # Tworzenie kursora
                cursor = connection.cursor()

                ORDER_NO = ORDER_no
                RELEASE_NO = RELEASE_no
                SEQUENCE_NO = SEQUENCE_no

                #SQL
                kwerenda = f"select CF$_PP_ETYKIETA1 from ifsapp.SHOP_ORD_CFV where ORDER_NO = '{ORDER_NO}' AND RELEASE_NO = '{RELEASE_NO}' AND SEQUENCE_NO = '{SEQUENCE_NO}'"

                """try:
                     cursor.execute(kwerenda)
                     for row in cursor:
                        print(row)
                except cx_Oracle.DatabaseError as e:
                    print("Błąd podczas wykonania zapytania:", e)"""
                cursor.execute(kwerenda)

                # pob info
                result = cursor.fetchone()
        
                #print(result)
                if result:
                    wynikk = result[0]
                    #print("PART_NO:", wynikk)
                    #etykieta = result.replace('\r\n', '\n')
                    wynikk = '\n'.join(el.strip() for el in wynikk.split('\n') if el.strip())
                    linie = wynikk.strip().split('\n')
                    poplinia = []
                    for linia in linie[:-1]:
                        if linia.strip() != "x":
                            poplinia.append(linia)
                        if linia.strip() == "x":
                            wzorwym = r'(\d+[x*]\d+[x*]\d+)'
                            wzorgat = r'/(\d{4})'
                            kwerendadod = f"select CF$_PP_OPIS from ifsapp.SHOP_ORD_CFV where ORDER_NO = '{ORDER_NO}' AND RELEASE_NO = '{RELEASE_NO}' AND SEQUENCE_NO = '{SEQUENCE_NO}'"
                            cursor.execute(kwerendadod)
                            #print(kwerendadod)
                            opis = cursor.fetchone()
                            #print(opis)
                            for wymiars in opis:
                                try:
                                    wymiarsur = re.findall(wzorwym, wymiars)
                                    for wym in wymiarsur: 
                                        #print(wym)
                                        poplinia.append(wym)
                                except Exception as e:
                                    print("Błąd przetwarzania:", e)
                            for gatuneks in opis:
                                try:
                                    gatuneksur = re.findall(wzorgat, gatuneks)
                                    for gat in gatuneksur:
                                        #print(gat)
                                        poplinia.append(gat)
                                except Exception as e:
                                    print("Błąd przetwarzania:", e)
                        etykieta = '\n'.join(poplinia)
                    return (etykieta)
                else:
                    print("Brak wyników dla podanych parametrów")
                
                # Zamknij oracla
                cursor.close()
                connection.close()

                
                #received_text = result #"PROPLASTICAPL-396*696*56/1730396x696x561730PN:13653/*/*2"
                #parts = received_text.split("PN:")
                #wymiar = parts[0].replace("PROPLASTICAPL-", "")
                #wymiar_do_pn = wymiar.split("*")
                #numer_pn = parts[1]
                """with open("wynik.txt", "w") as file:
                    file.write("PROPLASTICA\n")
                    file.write(wymiar + "\n")
                    file.write("*".join(wymiar_do_pn[:-1]) + "\n")
                    file.write(wymiar + "\n")
                    file.write(numer_pn + received_text + "\n")"""


                #print(lista)
                #print("- ####### -")

        #What if ther user press the interruption shortcut? 
        except KeyboardInterrupt:
            print("program został zakończony przez użytkownika.")
            return

        #if(len(lista)>0):
            #etykieta = '\n'.join(lista)
           #etykieta = '\n'.join(lista)
        if len(etykieta) < 10:
            #raise Exception("cause of the problem")
            return ("99#9")
        else:
                #PN = lista[len(lista)-1]
                #PN = PN[4:]
                #print(len(etykieta))
                #print('Wprowadziłeś etykietę:')
                #print(etykieta)
            print('\n')
            return   (etykieta)
    else:
        return ("99#9")

#===================================================================================================================================
def uruchom_aplikacje_i_nacisnij_klawisze(sciezka_pliku):

    """# Uruchamia plik przy użyciu domyślnej aplikacji
    #subprocess.Popen([nazwa_pliku])
    os.startfile(sciezka_pliku)
    # Czekaj na uruchomienie aplikacji i przydzielenie jej czasu na aktywację
    time.sleep(2)

    # Pobierz listę okien
    okna = gw.getWindowsWithTitle("EzCad2")
    print(okna)
    time.sleep(2)
    # Sprawdź, czy okno Notatnika jest aktywne
    for okno in okna:
        if "laser.ezd" in okno.title:
            # Ustaw focus na oknie aplikacji
            okno.activate()
        time.sleep(1)
        pyautogui.hotkey('alt', 'tab')"""
    """
        #pyautogui.click(900,508)
        #pyautogui.click(56,127)
	     pyautogui.click(890,475)
        print("pierwszy klik")
        time.sleep(2)
        pyautogui.doubleClick(85,720)
        print("drugi klik")
        time.sleep(2)
        #pyautogui.doubleClick(900,508)          

        #symuluj wklejenie tekstu etykiety
        pyautogui.typewrite(etykieta_global)
        time.sleep(2)
        pyautogui.hotkey('alt', 'o')
        time.sleep(2)
        pyautogui.hotkey('alt', 'tab')"""
    #try:
       # while True:
    time.sleep(1)
    print('Chcesz wypalić etykietę która jest bezpośrednio nad tym wierszem ? -> zeskanuj qr kod "Wypal",\nNie chcesz tej etykiety - zeskanuj qr kod "Inna etykieta" ') 
    
    
    klawisz1 = input().strip()
    klawisz1 = str(klawisz1)
    #print(str(klawisz1))
    time.sleep(1)
    #pyautogui.press(klawisz1)
    #klawisz1=jaki_klawisz()
    if '99#9' in klawisz1:
        print('błędny wybór')
    if 'Tak' in klawisz1:
        
        #time.sleep(1)
        # Uruchamia plik przy użyciu domyślnej aplikacji
        #subprocess.Popen([nazwa_pliku])
        os.startfile(sciezka_pliku)
        #print(sciezka_pliku)
        # Czekaj na uruchomienie aplikacji i przydzielenie jej czasu na aktywację
        time.sleep(2)

        # Pobierz listę okien
        okna = gw.getWindowsWithTitle("EzCad-Lite")
        #print(okna)
        time.sleep(3)
        # Sprawdź, czy okno Notatnika jest aktywne
        for okno in okna:
            if "laser.ezd" or "EzCad-Lite" in okno.title:
                # Ustaw focus na oknie aplikacji
                okno.activate()
                #print(okno)
                time.sleep(2)
                pyautogui.click(336,96)
            #pyautogui.hotkey('alt', 'tab')
            #pyautogui.hotkey('alt', 'tab')
                time.sleep(1)
                pyautogui.press('f2')
                time.sleep(16)
                #pyautogui.press('enter')
                #time.sleep(1)
                #print(okno)

                pyautogui.hotkey('alt', 'f4')
                #print(okno)
                time.sleep(0.3)
                pyautogui.press('n')
                return
    """if klawisz1=='2':
        print('Jeżeli wszystko jest gotowe do wypalenia, wpisz "y"')
        try:
            check = input()
            if len(check) > 0 and (check == 'y' or check.lower() == 'y'):
                pyautogui.hotkey('alt', 'tab')
                time.sleep(2)
                pyautogui.press('f2')
                time.sleep(10)
                pyautogui.hotkey('alt', 'f4')
                #break
        #What if ther user press the interruption shortcut? 
        except KeyboardInterrupt:
            print("program został zakończony przez użytkownika.")"""

    if 'inna' in klawisz1:
        time.sleep(0.5)
        #pyautogui.hotkey('alt', 'f4')
        """time.sleep(2)
        pyautogui.hotkey('alt', 'tab')
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'tab')"""
        
        #break
        #What if ther user press the interruption shortcut? 
    #except (ValueError, RuntimeError, TypeError, EOFError, KeyboardInterrupt):
        #print("program został zakończony.")        
    return


#===================================================================================================================================
def zamknij_okna():
    okna = gw.getWindowsWithTitle("EzCad-Lite")
    for okno in okna:
        #print(okno.title)
        if 'laser.ezd' in okno.title:
            okno.close()



#===================================================================================================================================
def jaki_klawisz():
    #wynik=''
    #print("Podaj klawisz")
    try:

        while True:
            wrd = input()
            if len(wrd) > 0 and str(wrd).isdigit:
                return (wrd)
            else:
                return ("99#9")

    #What if ther user press the interruption shortcut? 
    except KeyboardInterrupt:
        print("program został zakończony przez użytkownika.")
        return




#MAIN
#===================================================================================================================================
#pyautogui.hotkey('alt', 'tab')
try:
    while True:            
        stop=0
        etykieta_global = multi_input()
        #print(etykieta_global)

        if etykieta_global == "99#9":
            stop=1


        if stop==0:
            stop2=0
            zapisz_do_pliku(full_path_txt, str(etykieta_global))
            czy_plik=czy_plik_istnieje_i_aktualny(full_path_txt, 1)
            if czy_plik:
                print('plik istnieje i jest aktualny')
                print("Oto tekst do wypalenia:\r\n", str(etykieta_global).strip())
            else:
                stop2=1

            if stop2==0:
                #zamknij_okna()
                try:
                    while True:
                        #okna = gw.getWindowsWithTitle("znakowarka_EzCadA")
                        #print(okna)
                        time.sleep(1)
                        # Sprawdź, czy okno Notatnika jest aktywne
                        #for okno in okna:
                            #if "znakowarka_EzCadA" in okno.title:
                            # Ustaw focus na oknie aplikacji
                                #okno.activate()
                                #print(okno)
                        """print('zaznaczenie obszaru wypalania na płycie wpisz 1, wypalenie Laserem wpisz 2, koniec wpisz 3')
                        klawisz=jaki_klawisz()
                        if klawisz=="99#9":
                            print('błędny wybór')
                        if klawisz=='1':"""
                        uruchom_aplikacje_i_nacisnij_klawisze(full_path_ezd)
                        """ time.sleep(10)
                            pyautogui.press('enter')
                            time.sleep(2)
                            zamknij_okna()
                        if klawisz=='2':"""
                        """print('Jeżeli wszystko jest gotowe do wypalenia, wpisz "y"')
                            try:
                                check = input()
                                if len(check) > 0 and (check == 'y' or check.lower() == 'y'):
                                    uruchom_aplikacje_i_nacisnij_klawisze(full_path_ezd,domyslna_aplikacja,'f2')
                                    time.sleep(10)
                                    zamknij_okna()
                                    break
                            #What if ther user press the interruption shortcut? 
                            except KeyboardInterrupt:
                                print("program został zakończony przez użytkownika.")"""
                        """if klawisz=='3':
                            zamknij_okna()
                            break"""
                        os.system('cls' if os.name == 'nt' else 'clear')  # Wyczyść ekran konsoli

                        break
                #What if ther user press the interruption shortcut? 
                except (ValueError, RuntimeError, TypeError, EOFError, KeyboardInterrupt):
                    print("program został zakończony.")
        
            else:
                print('wczytaj ponownie etykiete, wystąpil problem z zapisem do pliku')
        








#What if ther user press the interruption shortcut? 
except (ValueError, RuntimeError, TypeError, EOFError, KeyboardInterrupt):
        print("program został zakończony.")