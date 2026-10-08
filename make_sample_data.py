"""Generate synthetic CSV exports in the layout consolidate.py expects.

Everything here is invented: document numbers, product codes, characteristic
keys and customer names. Run: python make_sample_data.py [folder]
"""
import csv
import random
import sys
from pathlib import Path

MODULES = ["MOD-R100", "MOD-R200", "MOD-B300", "MOD-C90", "MOD-M45"]
CUSTOMERS = ["Acme Logistics", "Northwind Parcel", "Globex Fulfilment", "Initech Distribution"]
CHARACTERISTICS = {
    "LEN_MM": [600, 1200, 1800, 2400, 3000],
    "WIDTH_MM": [420, 620, 840],
    "PITCH_MM": [60, 90, 120],
    "DRIVE_POWER_W": [35, 50],
    "DRIVE_VOLTAGE": [24, 48],
    "ZONE_LEN_MM": [600, 900, 1200],
    "CTRL_TYPE": ["Standalone", "Fieldbus"],
    "SIDE_GUIDE": ["None", "Fixed", "Adjustable"],
}


def write_export(folder, number, customer, rng):
    path = folder / f"export_{number}_{customer}.csv"
    with open(path, "w", newline="", encoding="latin1") as f:
        out = csv.writer(f, delimiter=";")
        out.writerow([number, "BoQ export"])
        for _ in range(rng.randint(2, 6)):
            delivery = f"{rng.randint(1, 28):02d}.{rng.randint(1, 12):02d}.2025"
            out.writerow(["", "", rng.choice(MODULES), rng.randint(1, 12), "", "", "", "", delivery])
            for key, values in CHARACTERISTICS.items():
                out.writerow(["", "", "", "", key, rng.choice(values)])
    return path


def main(folder="sample_data", orders=8, quotations=12, seed=7):
    rng = random.Random(seed)
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    for i in range(orders):
        write_export(folder, 5000100 + i, rng.choice(CUSTOMERS), rng)
    for i in range(quotations):
        write_export(folder, 2000100 + i, rng.choice(CUSTOMERS), rng)
    print(f"{orders + quotations} files written to {folder}")


if __name__ == "__main__":
    main(*sys.argv[1:2])
