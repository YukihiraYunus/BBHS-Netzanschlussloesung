from Strom.Strom import run as run_strom
from Gas.Gas import run as run_gas
from Wasser.Wasser import run as run_wasser
from Fernwärme.Fernwärme import run as run_fern
from Verbrauchsgeräte.Raumkühlung import run as run_raum
from Verbrauchsgeräte.Wärmepumpe import run as run_pumpe
from Verbrauchsgeräte.Stromspeicher import run as run_speicher
from Verbrauchsgeräte.Ladeeinrichtung import run as run_lade



# ======================
# Nur eins aktiv lassen!
# ======================

#run_strom()  
#run_gas()
#run_wasser()
#run_fern()   

#Verbrauchsgeräte:
#run_raum()
#run_pumpe()
#run_speicher()
run_lade()
