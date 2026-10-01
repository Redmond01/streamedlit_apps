from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


class ConfigError(ValueError):
    """Raised when bank lodgment configuration validation fails."""
    pass


DEFAULT_BRANCH_ID_MAP: dict[str, str] = {
    "BR_CARGO_1": "Cargo 1 OBX",
    "BR_JAH_MICHAEL": "Jah Michael TA",
    "BR_IPAJA": "Ipaja",
    "BR_BODIJA": "Bodija OBX",
    "BR_OGBOMOSO_1": "Ogbomoso 1",
    "BR_LEKKI_4": "Ajah",
    "BR_IFE_RD_3": "Ife 3 OBX",
    "BR_IFO": "Ifo SGR",
}

DEFAULT_PHONE_MAP: dict[str, str] = {
    "+2347082928596": "Iwo 1 OBX",
    "+2349115977488": "Cargo 2 OBX",
    "+2348057699971": "Sagam 4",
    "+2348169755874": "Lekki 3",
    "+2348064220606": "Ife 2 OBX",
    "+2348035016257": "Imota 1 SGR",
    "+2349076345569": "Lekki 1",
    "+2348130901883": "Oyo Ilora",
    "+2348103029619": "Itamaga SGR",
    "+2348060591404": "Igando SGR",
    "+2347039162599": "Guru OBX",
    "+2348032526950": "Sagam 1",
    "+2348038450077": "Jah Michael TA",
    "+2349030767371": "Ipaja",
    "+2347026483983": "Ifo SGR",
    "+2349069448010": "Ife 3 OBX",
    "+2347019606328": "Bodija OBX",
    "+2347080373413": "Ogbomoso 1",
    "+2347062939869": "Cargo 1 OBX",
    "+2348169435355": "Ajah",
    "+2348100682694": "Tollgate",
    "+2349018896657": "Ogbomoso 2",
    "+2347035351168": "Iwo 2 Monatan OBX",
    "+2347069038671": "Sagam 2",
    "+2348029929255": "Ife 1 OBX",
    "+2348166236126": "Olorunsogo",
    "+2347063250572": "Ijebu 2 SGR",
    "+2349044958279": "Ijebu 1 SGR",
    "+2347069557334": "Olomi 2",
    "+2347066364245": "Imota 2 SGR",
    "+2348068693767": "Olomi 1",
    "+2347037064054": "Apata SGR",
    "+2348138146855": "Sango",
    "+2348139253104": "Ogijo",
    "+2349155635023": "Agodi OBX",
}

DEFAULT_HANDLE_MAP: dict[str, str] = {
    "Duduke Mi": "Iwo 1 OBX",
    "Òĺámïļêķăñ": "Cargo 2 OBX",
    "Wealth": "Sagam 4",
    "Papa Tees Collections": "Oyo Ilora",
    "Bisola": "Itamaga SGR",
    "Bammy Pets": "Igando SGR",
    "ojumoolasayo7@gmail.com": "Guru OBX",
    "Olowoseunre Temitope": "Sagam 1",
    "Aperireeee": "Lekki 3",
    "olayemi65@gmail.com": "Jah Michael TA",
    "TEEDOSH": "Ipaja",
    "D Money": "Ife 3 OBX",
    "Q Quality": "Bodija OBX",
    "davidomenka68@gmail.com": "Ogbomoso 1",
    "Hon. Olusegun Edun": "Cargo 1 OBX",
    "Relentless": "Ajah",
    "Blessing Jonathan": "Tollgate",
    "PRINCE DEMOLA": "Ogbomoso 2",
    "Debo": "Iwo 2 Monatan OBX",
    "Kyle": "Sagam 2",
    "ABIMBOLA": "Ife 1 OBX",
    "Itunu": "Olorunsogo",
    "ADENIJI ADUNOLA": "Ijebu 2 SGR",
    "HR SGR": "Eleko 2",
    "mikekay0328": "Ijebu 1 SGR",
    "tenny4real1984": "Olomi 2",
    "Prince Arun": "Imota 2 SGR",
    "Donald Nwankwo": "Olomi 1",
    "Ebby": "Apata SGR",
    "Seyifly": "Sango",
    "Afolabi Mubarak": "Ogijo",
    "Akeem Olaitan Shittu": "Ife 2 OBX",
    "Babatunde Adeyemo": "Imota 1 SGR",
    "Omotayo": "Agodi OBX",
    "ibrahimolushola129": "Lekki 1",
}


@dataclass
class BankLodgmentConfig:
    target_date: str | None = None  # Expected as YYYY-MM-DD or DD-MM-YYYY
    auto_create_sheet: bool = True
    color_mode: str = "row"  # "row" (full row) or "columns" (cols G-I)
    flag_value: str = "C"
    branch_id_map: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_BRANCH_ID_MAP))
    phone_map: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_PHONE_MAP))
    handle_map: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_HANDLE_MAP))

    def validate(self) -> None:
        if self.target_date is not None and not self.target_date.strip():
            raise ConfigError("target_date cannot be empty string.")
        if self.color_mode not in {"row", "columns"}:
            raise ConfigError(f"Invalid color_mode: '{self.color_mode}'. Must be 'row' or 'columns'.")
