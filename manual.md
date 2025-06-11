# Anleitung Express
#### <span style="color: rgba(255,255,255,0.4)">Stand 10.09.2024</span>

## Teilautomatisierung der Transportprüfung

<style>
* {
    cursor: default;
}
a {
    color: white;
    transition: all 0.1s ease-in-out;
}

a:hover {
    color: rgb(150, 150, 255);
    text-decoration: none;
    font-weight: bold;
}
</style>

### Inhalt
* Module und Bibliotheken
    * PyQt6.QtWidgets
        * <a title="Verwaltet anwendungsweite Ressourcen und Einstellungen.">QApplication</a>
        * <a title="Zeigt einen Text oder ein Bild an.">QLabel</a>
        * <a title="Basisklasse für alle UI-Objekte in PyQt.">QWidget</a>
        * <a title="Stellt eine Schaltfläche dar.">QPushButton</a>
        * <a title="Zeigt eine Liste von Elementen an.">QListWidget</a>
        * <a title="Stellt ein Element in einer QListWidget dar.">QListWidgetItem</a>
        * <a title="Einzeiliger Texteditor.">QLineEdit</a>
        * <a title="Zeigt Daten in Tabellenform an.">QTableWidget</a>
        * <a title="Stellt ein Element in einem QTableWidget dar.">QTableWidgetItem</a>
        * <a title="Verwaltet die Kopfzeilen einer Tabelle.">QHeaderView</a>
        * <a title="Dropdown-Liste zur Auswahl von Elementen.">QComboBox</a>
    * PyQt6.QtCore
        * <a title="Kernfunktionalität ohne GUI.">Qt</a>
        * <a title="Definiert die Größe von Widgets.">QSize</a>
        * <a title="Verwaltet Eigenschaftsanimationen.">QPropertyAnimation</a>
        * <a title="Definiert die Geometrie von Widgets.">QRect</a>
        * <a title="Bietet Easing-Kurven für Animationen.">QEasingCurve</a>
    * PyQt6.QtGui
        * <a title="Verwaltet Symbole und Bilder.">QIcon</a>
    * datetime
        * <a title="Verwaltet Datums- und Zeitoperationen.">datetime (dt)</a>
        * <a title="Stellt den Unterschied zwischen zwei Daten dar.">timedelta</a>
    * <a title="Bietet COM-Client-Unterstützung.">win32com.client (client)</a>
    * tkinter
        * <a title="Öffnet Dateidialoge.">filedialog</a>
    * <a title="Datenanalysebibliothek.">polars (pl)</a>
    * <a title="Numerische Berechnungen in Python.">numpy (np)</a>
    * <a title="Python COM-Schnittstelle.">pythoncom</a>
    * <a title="Liest und schreibt Excel-Dateien.">openpyxl</a>
    * <a title="Erkennt die Kodierung von Textdateien.">chardet</a>
    * <a title="Bietet Funktionen zur Interaktion mit dem Betriebssystem.">os</a>
    * <a title="Modul für reguläre Ausdrücke.">re</a>

    