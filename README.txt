TTC – Evidență Atestate și Inspecții (Windows)

Proiectul este pregătit pentru generarea aplicației Windows EXE prin GitHub Actions.

Pași:
1. Creează un cont GitHub (dacă nu ai deja).
2. Creează un repository nou, de exemplu TTC-Evidenta.
3. Încarcă toate fișierele și directoarele din acest ZIP în repository.
4. În repository, deschide Actions.
5. Dacă este necesar, activează Actions și permite workflow-urile.
6. Workflow-ul Build TTC Windows app se va executa la push; alternativ, din Actions selectează workflow-ul și apasă Run workflow.
7. După build, deschide execuția terminată și descarcă artifact-ul TTC-Evidenta-Documente-Windows.
8. Extrage artifact-ul și pornește EXE-ul.

Datele aplicației sunt stocate local în %APPDATA%\EvidentaDocumente\evidenta.db, compatibil cu baza versiunii anterioare, pentru a păstra înregistrările.

Coloane: Nr. referință (ID automat), Data înregistrării, Nume, Prenume, Telefon, Serviciu, Nr. auto, Data efectuării, Data expirării și Status.
Data înregistrării se completează automat și poate fi ajustată.
