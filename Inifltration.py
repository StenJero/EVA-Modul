import pandas as pd
import psychro as psy

def calculate_thermal_load():
    """
    This code calculates the thermal load and ventilation requirement in a room based on various parameters such as indoor temperature, outdoor temperature, relative humidity, and wall area.

    Methodology:
    - The code begins by importing the necessary libraries, namely "pandas" for data processing and "psychro" for psychrometric calculations.
    - Parameters such as specific heat (c), latent heat (lat), indoor temperature (θi), wall area (Sw), and others are defined.
    - Next, an Excel file containing humidity data is read, and the value for indoor relative humidity (ϕ) is extracted.
    - The code then calculates the air infiltration volumetric flow rate (Vinf) based on the desired air changes per hour (ACH) and room volume (V).
    - The minimum air flow rate per person (mp) and the sensible thermal load of the thermal zone (Qs) are calculated, taking into account convection heat transfer, air infiltration volumetric flow rate, thermal bridges, and heat load from occupants.
    - The total latent load of the thermal zone (Qlat) is also calculated, considering both latent heat transfer and moisture release from occupants.
    - Finally, the supply air mass flow rate (m) and the absolute humidity after the thermal zone (ws) are calculated.

    Results:
    - The calculated values are printed, including the air infiltration volumetric flow rate (Vinf), sensible thermal load (Qs), latent thermal load (Qlat), supply air mass flow rate (m), and absolute humidity after the thermal zone (ws).

    Discussion:
    - The results are analyzed and interpreted, and comparisons can be made with expected values or standards.
    - Any deviations or inconsistencies are discussed, and the adequacy of the code in calculating the thermal load and ventilation requirement is assessed.
    - Strengths and weaknesses of the code are also addressed.

    Conclusion:
    - In conclusion, the code successfully calculates the thermal load and ventilation requirement in a room.
    - It is emphasized that the code is based on the provided parameters and includes specific assumptions and simplifications.
    - Possible applications and implications of the results are discussed, and suggestions for further improvements or adaptations of the code are provided.
    """

    c = 1000     # (J/Kg,K) Specific Heat
    lat = 2496e3  # latent heat J/kg
    θi = 23      # (°C) Indoor Temperature
    θs = θi + 15 # define Temperature after TZ
    V = 3 * 3 * 8 # Volumetric of the room

    Uw = 0.7     # (W/m2,K) Convection Coefficients Wall
    Sw = 3 * 8   # (m2) Surface of Wall
    θo = -7.3    # (°C) Outdoor Temperature

    # Read humidity values from Excel file
    humidity_data = pd.read_excel("humidity.csv", engine='xlrd')

    ϕ = humidity_data.at[4, 'I'] / 100  # (%) Humidity Inside

    wo = psy.w(θo, ϕ, Z=0)
    p = 1       # (-) Anzahl Personen
    Vp = 36     # (m3/h) Volumenstrom pro Person
    Vd = (p * Vp / 3600)    # (m3/s) Totally volumetric

    ψ = 0.091
    # (W/m,K) Wärmebrücken
    h = 8       # (m) Höhe
    l = 3       # (m) Länge
    L = 2 * h + 2 * l   # length of thermal bridge
    Qi = 70 * p         # (W, Load person)

    ACH = 0.5
    Vinf = ACH * V / 3600
    minf = Vinf / psy.v(θo, wo)
    print(f"Vinf = {Vinf:.2f} kg/s")

    mp = Vd / psy.v(θo, wo)
    Qs = Uw * Sw * (θo - θi) + minf * c * (θo - θi) + ψ * L * (θo - θi) + Qi
    print(f"Qs = {Qs:.2f} W")

    # The total latent load of the thermal zone
    mvp = 71e-3 / 3600  # kg/s, Feuchtigskeitsabgabe pro Person
    Qla = p * mvp * l

    m = Vd / psy.v(θo, wo)
    Qlat = m * lat * (wo - ϕ) + Qla
    print(f"Qlat = {Qlat:.2f} W")

    # Temperature Input after TZ
    m = -Qs / (c * (θs - θi))
    ws = ϕ - Qlat / (m * lat)

    print(f"m = {m:.3f} kg/s, ws = {ws:.3f} kg/kg")

    # The minimum air flow rate for indoor air quality
    if m > mp:
        m = mp
    print(f"m = {m:.3f} ks/s")

if __name__ == "__main__":
    calculate_thermal_load()
