from pyautocad import Autocad, APoint
import pandas as pd


def teken_kruizen(acad, aantal_kruizen, afstand_tussen_kruizen):
    template_laag = acad.doc.Layers.Add("template")
    template_laag.color = 1  # Rood

    # Stel de huidige laag in op 'template'
    acad.doc.ActiveLayer = template_laag

    start_x = 0
    for i in range(aantal_kruizen):
        middenpunt = APoint(start_x + i * afstand_tussen_kruizen, 0)
        # Teken horizontale lijn (50 mm aan beide zijden van het middenpunt)
        acad.model.AddLine(middenpunt + APoint(-50, 0), middenpunt + APoint(50, 0))
        # Teken verticale lijn (50 mm aan beide zijden van het middenpunt)
        acad.model.AddLine(middenpunt + APoint(0, -50), middenpunt + APoint(0, 50))


def selecteer_basis_laag(acad):
    """Selecteert de 'Basis' laag of creëert deze als hij niet bestaat."""
    basis_laag_naam = "Basis"
    if basis_laag_naam not in [laag.Name for laag in acad.doc.Layers]:
        acad.doc.Layers.Add(basis_laag_naam)
    acad.doc.ActiveLayer = acad.doc.Layers.Item(basis_laag_naam)
    print(f"Actieve laag teruggezet naar '{basis_laag_naam}'.")


def classificeer_lijnen_en_verwijder_dubbelen(acad, afstand_tussen_kruizen, aantal_kruizen):
    frames_data = {}
    unieke_segmenten = set()

    # Voorbereiding van frame datastructuur
    for i in range(1, aantal_kruizen + 1):
        frames_data[f"Frame {i}"] = []

    for obj in acad.model:
        if obj.ObjectName == "AcDbLine" and obj.Layer != "template":
            frame_index = int(min(obj.StartPoint[0], obj.EndPoint[0]) // afstand_tussen_kruizen) + 1
            frame_x_start = (frame_index - 1) * afstand_tussen_kruizen

            # Coördinaten relatief aan frame beginpunt
            rel_start = (round(obj.StartPoint[0] - frame_x_start, 2), round(obj.StartPoint[1], 2))
            rel_end = (round(obj.EndPoint[0] - frame_x_start, 2), round(obj.EndPoint[1], 2))

            segment = (rel_start, rel_end) if rel_start < rel_end else (rel_end, rel_start)

            # Voeg segment toe als het uniek is
            if segment not in unieke_segmenten:
                unieke_segmenten.add(segment)
                frames_data[f"Frame {frame_index}"].append(segment)

    # Printen van de segmenten per frame, zonder dubbele coördinaten
    for frame, segmenten in frames_data.items():
        print(f"{frame} heeft {len(segmenten)} unieke segmenten:")
        for seg in segmenten:
            print(f"  Segment: Startpunt: X = {seg[0][0]}, Y = {seg[0][1]}, Eindpunt: X = {seg[1][0]}, Y = {seg[1][1]}")
    return frames_data


def haal_lijn_coordinaten(acad):
    lijnen_coordinaten = []

    for obj in acad.model:
        if obj.ObjectName == "AcDbLine":
            startpunt = (obj.StartPoint[0], obj.StartPoint[1])  # Geen offset correctie nodig
            eindpunt = (obj.EndPoint[0], obj.EndPoint[1])
            lijnen_coordinaten.append((startpunt, eindpunt))
    
    # Extraheren en sorteren van unieke coördinaten
    unieke_coordinaten = set()
    for lijn in lijnen_coordinaten:
        unieke_coordinaten.update([lijn[0], lijn[1]])
    gesorteerde_coordinaten = sorted(unieke_coordinaten, key=lambda coord: coord[1])

    for coord in gesorteerde_coordinaten:
        print(f"Coördinaat: X = {coord[0]}, Y = {coord[1]}")


def formatteer_lijn_coordinaten(frames_data):
    for frame, segmenten in frames_data.items():
        print(f"{frame}:")
        for index, segment in enumerate(segmenten):
            # Toon startpunt voor elke lijn
            print(f"{segment[0][0]},{segment[0][1]}")
            
            # Voor de laatste lijn in de reeks, toon ook het eindpunt
            if index == len(segmenten) - 1:
                print(f"{segment[1][0]},{segment[1][1]}")


def opslaan_in_excel(frames_data, bestandsnaam="CoordinatenFrames.xlsx"):
    with pd.ExcelWriter(bestandsnaam, engine='openpyxl') as writer:
        for frame, segmenten in frames_data.items():
            # Verzamel alle punten (inclusief start- en eindpunten van het laatste segment)
            punten = [seg[0] for seg in segmenten]  # Voeg alle startpunten toe
            if segmenten:  # Voeg eindpunt van het laatste segment toe
                punten.append(segmenten[-1][1])

            # Convert de punten van mm naar m
            punten = [(x/1000, y/1000) for x, y in punten]

            # Maak een DataFrame van de punten
            df = pd.DataFrame(punten, columns=['X (m)', 'Y (m)'])
            
            # Sorteer de DataFrame op Y-waarde
            df_sorted = df.sort_values(by='Y (m)')
            
            # Schrijf de gesorteerde DataFrame naar een blad in het Excel-bestand
            df_sorted.to_excel(writer, sheet_name=frame, index=False)

    print(f"Data opgeslagen in {bestandsnaam} in meters.")


def main():
    acad = Autocad(create_if_not_exists=True)
    aantal_kruizen = int(input("Voer het aantal kruizen in: "))
    afstand_tussen_kruizen = float(input("Voer de afstand tussen de kruizen in (mm): "))
    
    teken_kruizen(acad, aantal_kruizen, afstand_tussen_kruizen)

    selecteer_basis_laag(acad)

    print("Teken nu de lijnen op de kruizen. Druk op Enter wanneer je klaar bent...")
    input()

    # Classificeer de getekende lijnen per frame
    classificeer_lijnen_en_verwijder_dubbelen(acad, afstand_tussen_kruizen, aantal_kruizen)

    # Aangenomen dat `frames_data` je georganiseerde lijnsegmenten bevat
    frames_data = classificeer_lijnen_en_verwijder_dubbelen(acad, afstand_tussen_kruizen, aantal_kruizen)
    formatteer_lijn_coordinaten(frames_data)

    opslaan_in_excel(frames_data)

if __name__ == "__main__":
    main()