import os
import shutil
import re
from collections import defaultdict

def extract_nume_prenume(filename):
    """
    Extrage numele și prenumele din numele fișierului.
    Presupune formatul: "NUME PRENUME orice_altceva.extensie"
    """
    # Elimină extensia fișierului
    base_name = os.path.splitext(filename)[0]
    
    # Pattern pentru caractere românești: include toate variantele de diacritice
    # ş și ș (ambele variante), ţ și ț (ambele variante)
    romanian_pattern = r'^([A-Za-zÀ-ÿăâîșțĂÂÎȘȚşţŞŢ]+)\s+([A-Za-zÀ-ÿăâîșțĂÂÎȘȚşţŞŢ]+)'
    match = re.match(romanian_pattern, base_name, re.UNICODE)
    
    if match:
        nume = match.group(1).strip()
        prenume = match.group(2).strip()
        return nume + " " + prenume
    
    # Fallback: încearcă să găsească orice două cuvinte la început
    fallback_match = re.match(r'^(\S+)\s+(\S+)', base_name)
    if fallback_match:
        nume = fallback_match.group(1).strip()
        prenume = fallback_match.group(2).strip()
        # Verifică dacă conțin doar caractere valide românești
        check_pattern = r'^[A-Za-zÀ-ÿăâîșțĂÂÎȘȚşţŞŢ]+$'
        if re.match(check_pattern, nume) and re.match(check_pattern, prenume):
            return nume + " " + prenume
    
    return None

def gaseste_toate_directoarele(director_curent=".", nivel_max=2):
    """
    Găsește toate directoarele cu fișiere .docx din directorul curent
    Caută până la nivel_max nivele în adâncime
    """
    directoare_cu_docx = []
    
    def cauta_recursive(cale, nivel_curent=0):
        if nivel_curent > nivel_max:
            return
            
        try:
            items = os.listdir(cale)
        except PermissionError:
            return
        
        # Verifică fișierele direct în acest director
        fisiere_docx = []
        subdirectoare = []
        
        for item in items:
            cale_completa = os.path.join(cale, item)
            
            if os.path.isfile(cale_completa):
                if item.lower().endswith('.docx') and not item.startswith('~'):
                    fisiere_docx.append(item)
            elif os.path.isdir(cale_completa) and item != "FISIER ORGANIZAT":
                subdirectoare.append((item, cale_completa))
        
        # Dacă acest director conține fișiere .docx, adaugă-l
        if fisiere_docx:
            cale_relativa = os.path.relpath(cale, director_curent)
            directoare_cu_docx.append(cale_relativa)
            nivel_info = f" (nivel {nivel_curent})" if nivel_curent > 0 else ""
            print(f"📁 Găsit director cu {len(fisiere_docx)} fișiere .docx: {cale_relativa}{nivel_info}")
        
        # Caută în subdirectoare
        for nume_subdir, cale_subdir in subdirectoare:
            cauta_recursive(cale_subdir, nivel_curent + 1)
    
    print("🔍 Caut directoare cu fișiere .docx (până la 2 nivele adâncime)...")
    cauta_recursive(director_curent, 0)
    
    return directoare_cu_docx

def organizeaza_contracte_automat():
    """
    Organizează contractele din TOATE directoarele găsite automat
    în folderul FISIER ORGANIZAT, grupate pe persoane.
    """
    
    # Găsește automat toate directoarele cu fișiere .docx
    directoare_sursa = gaseste_toate_directoarele()
    
    if not directoare_sursa:
        print("❌ Nu s-au găsit directoare cu fișiere .docx!")
        return
    
    print(f"\n🔍 Găsite {len(directoare_sursa)} directoare cu documente:")
    for i, director in enumerate(directoare_sursa, 1):
        print(f"  {i}. {director}")
    
    director_destinatie = 'FISIER ORGANIZAT'
    
    # Creează directorul de destinație dacă nu există
    if not os.path.exists(director_destinatie):
        os.makedirs(director_destinatie)
        print(f"\n✅ Creat director: {director_destinatie}")
    
    # Dicționar pentru a grupa fișierele pe persoane
    persoane_fisiere = defaultdict(list)
    total_fisiere = 0
    
    # Scanează toate directoarele găsite
    for director in directoare_sursa:
        director_path = director if director != "." else os.getcwd()
        director_name = "directorul curent" if director == "." else director
        
        if not os.path.exists(director) and director != ".":
            print(f"⚠️  Directorul '{director}' nu există - sărit")
            continue
            
        print(f"\n📁 Procesez {director_name}")
        fisiere_director = 0
        
        # Scanează fișierele din director
        fisiere_lista = os.listdir(director) if director != "." else os.listdir(".")
        
        for filename in fisiere_lista:
            filepath = os.path.join(director, filename) if director != "." else filename
            
            # Ignoră subdirectoarele și fișierele temporare
            if os.path.isdir(filepath) or filename.startswith('~'):
                continue
            
            # Procesează doar fișierele .docx
            if not filename.lower().endswith('.docx'):
                continue
                
            # Extrage numele și prenumele
            nume_prenume = extract_nume_prenume(filename)
            
            if nume_prenume:
                persoane_fisiere[nume_prenume].append({
                    'fisier': filename,
                    'director_sursa': director_name,
                    'cale_completa': filepath
                })
                print(f"  ✅ {nume_prenume} - {filename}")
                fisiere_director += 1
                total_fisiere += 1
            else:
                print(f"  ❌ Nu s-a putut extrage numele din: {filename}")
        
        print(f"  📊 Găsite {fisiere_director} contracte în {director_name}")
    
    if total_fisiere == 0:
        print("\n❌ Nu s-au găsit fișiere cu nume valid pentru organizare!")
        return
    
    print(f"\n📊 TOTAL: {total_fisiere} contracte găsite pentru {len(persoane_fisiere)} persoane")
    
    # Creează foldere și copiază fișierele
    contracte_copiate = 0
    for nume_prenume, fisiere_info in persoane_fisiere.items():
        # Creează folderul pentru persoană
        folder_persoana = os.path.join(director_destinatie, nume_prenume)
        
        if not os.path.exists(folder_persoana):
            os.makedirs(folder_persoana)
            print(f"\n📂 Creat folder: {nume_prenume}")
        
        # Copiază toate fișierele pentru această persoană
        for info in fisiere_info:
            sursa = info['cale_completa']
            destinatie = os.path.join(folder_persoana, info['fisier'])
            
            try:
                shutil.copy2(sursa, destinatie)
                print(f"  ✅ Copiat din {info['director_sursa']}: {info['fisier']}")
                contracte_copiate += 1
            except Exception as e:
                print(f"  ❌ Eroare la copierea {info['fisier']}: {str(e)}")
    
    # Afișează sumar final
    print(f"\n🎉 ORGANIZAREA COMPLETĂ!")
    print(f"📊 Persoane procesate: {len(persoane_fisiere)}")
    print(f"📄 Contracte copiate: {contracte_copiate}/{total_fisiere}")
    print(f"📁 Foldere create în '{director_destinatie}':")
    
    for nume_prenume, fisiere_info in persoane_fisiere.items():
        contracte_pe_tip = defaultdict(int)
        for info in fisiere_info:
            contracte_pe_tip[info['director_sursa']] += 1
        
        detalii = ", ".join([f"{tip}({nr})" for tip, nr in contracte_pe_tip.items()])
        print(f"  👤 {nume_prenume}: {len(fisiere_info)} contracte [{detalii}]")

def organizeaza_contracte_manual():
    """
    Organizează contractele din directoare specificate manual de utilizator
    """
    print("📝 Introdu directoarele sursă (unul pe linie, linie goală pentru a termina):")
    directoare_sursa = []
    
    while True:
        director = input("Director: ").strip()
        if not director:
            break
        
        if os.path.exists(director) and os.path.isdir(director):
            # Verifică dacă conține fișiere .docx
            fisiere_docx = [f for f in os.listdir(director) 
                           if f.lower().endswith('.docx') and not f.startswith('~')]
            
            if fisiere_docx:
                directoare_sursa.append(director)
                print(f"  ✅ Adăugat: {director} ({len(fisiere_docx)} fișiere .docx)")
            else:
                print(f"  ⚠️  {director} nu conține fișiere .docx")
        else:
            print(f"  ❌ Directorul nu există: {director}")
    
    if not directoare_sursa:
        print("❌ Nu s-au specificat directoare valide!")
        return
    
    # Continuă cu organizarea folosind directoarele specificate
    organizeaza_cu_directoare_specifice(directoare_sursa)

def organizeaza_cu_directoare_specifice(directoare_sursa):
    """
    Organizează contractele din directoare specifice
    """
    director_destinatie = 'FISIER ORGANIZAT'
    
    # Creează directorul de destinație dacă nu există
    if not os.path.exists(director_destinatie):
        os.makedirs(director_destinatie)
        print(f"✅ Creat director: {director_destinatie}")
    
    # Dicționar pentru a grupa fișierele pe persoane
    persoane_fisiere = defaultdict(list)
    total_fisiere = 0
    
    # Scanează toate directoarele specificate
    for director in directoare_sursa:
        print(f"\n📁 Procesez directorul: {director}")
        fisiere_director = 0
        
        # Scanează fișierele din director
        for filename in os.listdir(director):
            filepath = os.path.join(director, filename)
            
            # Ignoră subdirectoarele și fișierele temporare
            if os.path.isdir(filepath) or filename.startswith('~'):
                continue
            
            # Procesează doar fișierele .docx
            if not filename.lower().endswith('.docx'):
                continue
                
            # Extrage numele și prenumele
            nume_prenume = extract_nume_prenume(filename)
            
            if nume_prenume:
                persoane_fisiere[nume_prenume].append({
                    'fisier': filename,
                    'director_sursa': director,
                    'cale_completa': filepath
                })
                print(f"  ✅ {nume_prenume} - {filename}")
                fisiere_director += 1
                total_fisiere += 1
            else:
                print(f"  ❌ Nu s-a putut extrage numele din: {filename}")
        
        print(f"  📊 Găsite {fisiere_director} contracte în {director}")
    
    if total_fisiere == 0:
        print("\n❌ Nu s-au găsit fișiere cu nume valid pentru organizare!")
        return
    
    print(f"\n📊 TOTAL: {total_fisiere} contracte găsite pentru {len(persoane_fisiere)} persoane")
    
    # Creează foldere și copiază fișierele
    contracte_copiate = 0
    for nume_prenume, fisiere_info in persoane_fisiere.items():
        # Creează folderul pentru persoană
        folder_persoana = os.path.join(director_destinatie, nume_prenume)
        
        if not os.path.exists(folder_persoana):
            os.makedirs(folder_persoana)
            print(f"\n📂 Creat folder: {nume_prenume}")
        
        # Copiază toate fișierele pentru această persoană
        for info in fisiere_info:
            sursa = info['cale_completa']
            destinatie = os.path.join(folder_persoana, info['fisier'])
            
            try:
                shutil.copy2(sursa, destinatie)
                print(f"  ✅ Copiat din {info['director_sursa']}: {info['fisier']}")
                contracte_copiate += 1
            except Exception as e:
                print(f"  ❌ Eroare la copierea {info['fisier']}: {str(e)}")
    
    # Afișează sumar final
    print(f"\n🎉 ORGANIZAREA COMPLETĂ!")
    print(f"📊 Persoane procesate: {len(persoane_fisiere)}")
    print(f"📄 Contracte copiate: {contracte_copiate}/{total_fisiere}")
    print(f"📁 Foldere create în '{director_destinatie}':")
    
    for nume_prenume, fisiere_info in persoane_fisiere.items():
        contracte_pe_tip = defaultdict(int)
        for info in fisiere_info:
            contracte_pe_tip[info['director_sursa']] += 1
        
        detalii = ", ".join([f"{tip}({nr})" for tip, nr in contracte_pe_tip.items()])
        print(f"  👤 {nume_prenume}: {len(fisiere_info)} contracte [{detalii}]")

def main():
    print("=== ORGANIZATOR FLEXIBIL CONTRACTE ===\n")
    
    # Debug: afișează structura de directoare
    print("🔍 DEBUG - Structura directoarelor:")
    
    def afiseaza_structura(cale, prefix="", nivel=0):
        if nivel > 2:  # Limitează afișarea la 2 nivele
            return
            
        try:
            items = sorted(os.listdir(cale))
            directoare = []
            fisiere_docx = []
            
            for item in items:
                cale_item = os.path.join(cale, item)
                if os.path.isdir(cale_item) and item != "FISIER ORGANIZAT":
                    directoare.append(item)
                elif item.lower().endswith('.docx') and not item.startswith('~'):
                    fisiere_docx.append(item)
            
            # Afișează fișierele .docx din directorul curent
            for fisier in fisiere_docx:
                print(f"{prefix}📄 {fisier}")
            
            # Afișează subdirectoarele și conținutul lor
            for i, director in enumerate(directoare):
                is_last = i == len(directoare) - 1
                print(f"{prefix}📁 {director}/")
                
                # Afișează conținutul subdirectorului
                cale_subdir = os.path.join(cale, director)
                nou_prefix = prefix + ("    " if is_last else "│   ")
                afiseaza_structura(cale_subdir, nou_prefix, nivel + 1)
                
        except PermissionError:
            print(f"{prefix}❌ Acces interzis")
    
    afiseaza_structura(".")
    
    print("\n" + "="*50)
    print("Alege modul de lucru:")
    print("1. Automat - caută toate directoarele cu fișiere .docx")
    print("2. Manual - specifică directoarele sursă")
    
    alegere = input("\nIntroduci opțiunea (1 sau 2): ").strip()
    
    if alegere == "1":
        print("\n🔍 Mod automat - caut directoare cu fișiere .docx...")
        directoare_gasite = gaseste_toate_directoarele()
        
        if not directoare_gasite:
            print("❌ Nu s-au găsit directoare cu fișiere .docx!")
            print("💡 Încearcă modul manual (opțiunea 2)")
            return
        
        print(f"\n📋 Voi procesa {len(directoare_gasite)} directoare:")
        for director in directoare_gasite:
            if director == ".":
                print(f"  - Directorul curent")
            else:
                print(f"  - {director}")
        
        confirmare = input(f"\nContinui cu organizarea? (Enter pentru Da): ").strip()
        if confirmare and confirmare.lower() not in ['da', 'yes', 'y', '']:
            print("❌ Anulat.")
            return
        
        print("\n" + "="*60)
        organizeaza_contracte_automat()
        print("="*60)
        
    elif alegere == "2":
        print("\n📝 Mod manual - specifică directoarele...")
        organizeaza_contracte_manual()
        
    else:
        print("❌ Opțiune invalidă!")
        return
    
    print("\n✨ Gata! Verifică folderul 'FISIER ORGANIZAT'")
    input("Apasă Enter pentru a închide...")

if __name__ == "__main__":
    main()
