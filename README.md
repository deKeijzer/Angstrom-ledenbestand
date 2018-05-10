# Ångström ledenbestand
[![Binder](https://mybinder.org/badge.svg)](https://mybinder.org/v2/gh/deKeijzer/Angstrom-ledenbestand/master?urlpath=https%3A%2F%2Fgithub.com%2FdeKeijzer%2FAngstrom-ledenbestand%2Fblob%2Fmaster%2Fnotebook.ipynb)
Dit is de Python code welke gebruikt voor het ledenbestand van Ångström. Middels deze code kunnen de gegevens van mensen welke een Angfo en/of de nieuwsbrief willen ontvangen worden geexporteerd uit het ledenbestand. Meer opties zijn ook mogelijk wanneer de code word aangepast, dit is simpel te doen. Denk hierbij aan het opvragen van de gegevens van alle ereleden. De notebook is gebruikt om van het oude ledebestand (3 Excel sheets) te migreren naar het nieuwe ledenbestand, welke slechts 1 Excel sheet is.

# Gebruik
Gebruik de ['Clone or download'](https://github.com/deKeijzer/Angstrom-ledenbestand/archive/master.zip) knop om de code te downloaden. Laat alles in dezelfde map staan.
Over het algemeen is het gebruik van de Jupyter Notebook niet nodig. Simpelweg het runnen van [main.py](main.py) met Python is voldoende. Voor het gebruik van Jupyter Notebook run het bestand [run_notebook.bat](run_notebook.bat).
* Zorg ervoor dat 'Ledenbestand.xlsx' in de map 'input' staat.
* De code is gevoelig voor hoofdletters, let hier op.
Verander ook niet zomaar kolommen van naam in het ledenbestand.
Na het runnen van [main.py](main.py) zullen de gegevens van de mensen die de Ångfo en nieuwsbrief willen ontvangen zijn geëxporteerd naar de map 'output'.

# Extra informatie m.b.t. Jupyter Notebooks
Sommige code zal in de vorm van een .py bestand zijn. Echter komen er ook [Jupyter Notebooks](http://jupyter.org) voor, daar waar dat handiger is. Lees [hier](http://jupyter.readthedocs.io/en/latest/install.html) hoe je Jupyter installeert op je computer. Een korte uitleg: \
Installeer [Python3](https://www.python.org/). Open (op Windows) [CMD](https://en.wikipedia.org/wiki/Cmd.exe) als administrator (navigator? op macOS) en typ het volgende: \
`pip3 install jupyter` \
En voor bijvoorbeeld libraries: \
`pip3 install pandas` \
`pip3 install numpy` \
Om de notebook de starten kan `jupyter notebook` getypt worden, of run het run_notebook.bat bestand in de folder van de betreffende notebook.