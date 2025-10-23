# :earth_americas: GDP dashboard template

A simple Streamlit app showing the GDP of different countries in the world.

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://gdp-dashboard-template.streamlit.app/)

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```
def load_product_alignment():

    try:
        product_alignment = {
            "annuity": [
                "Bajaj Allianz Life Guaranteed Pension Goal II",
                "Bajaj Allianz Life Saral Pension"
            ],
            "combi": [
                "Bajaj Allianz Life Capital Goal Suraksha"
            ],
            "group": [
                "Bajaj Allianz Life Group Term Life",
                "Bajaj Allianz Life Group Credit Protection Plus",
                "Bajaj Allianz Life Group Sampoorna Jeevan Suraksha",
                "Bajaj Allianz Life Group Employee Benefit",
                "Bajaj Allianz Life Group Superannuation Secure Plus",
                "Bajaj Allianz Life Group Superannuation Secure",
                "Bajaj Allianz Life Group Employee Care",
                "Bajaj Allianz Life Group Secure Return",
                "Bajaj Allianz Life Group Sampoorna Suraksha Kavach",
                "Bajaj Allianz Life Pradhan Mantri Jeevan Jyoti Bima Yojana",
                "Bajaj Allianz Life Group Secure Shield",
                "Bajaj Allianz Life Group Investment Plan"
            ],
            "non_par": [
                "Bajaj Allianz Life Goal Suraksha",
                "Bajaj Allianz Life Assured Wealth Goal Platinum",
                "Bajaj Allianz Life Guaranteed Wealth Goal",
                "Bajaj Allianz Life Guaranteed Saving Goal",
                "Bajaj Allianz Life Assured Wealth Goal"
            ],
            "par": [
                "Bajaj Allianz Life ACE",
                "Bajaj Allianz Life ACE Advantage"
            ],
            "rider": [
                "Bajaj Allianz Accidental Death Benefit Rider",
                "Bajaj Allianz Accidental Permanent Total/Partial Disability Benefit Rider",
                "Bajaj Allianz Life Linked Accident Protection Rider II",
                "Bajaj Allianz Life Family Protect Rider",
                "Bajaj Allianz Life Group New Terminal Illness Rider",
                "Bajaj Allianz Life Group Accelerated Critical Illness Rider",
                "Bajaj Allianz Life Group Accidental Permanent Total/Partial Disability Benefit Rider",
                "Bajaj Allianz Life Group Critical Illness Rider",
                "Bajaj Allianz Life Group Accidental Death Benefit",
                "Bajaj Allianz Life New Critical Illness Benefit Rider",
                "Bajaj Allianz Life Care Plus Rider",
                "Bajaj Allianz Life Linked Critical Illness Benefit Rider"
            ],
            "term": [
                "Bajaj Allianz Life iSecure II",
                "Bajaj Allianz Life eTouch II",
                "Bajaj Allianz Life Saral Jeevan Bima",
                "Bajaj Allianz Life Diabetic Term Plan II Sub 8 HbA1c",
                "Bajaj Allianz Life Smart Protection Goal"
            ],
            "ulip": [
                "Bajaj Allianz Life Goal Assure IV",
                "Bajaj Allianz Life Magnum Fortune Plus III",
                "Bajaj Allianz Life Invest Protect Goal III",
                "Bajaj Allianz Life Fortune Gain II",
                "Bajaj Allianz Life Future Wealth Gain IV",
                "Bajaj Allianz Life LongLife Goal III",
                "Bajaj Allianz Life Smart Wealth Goal V",
                "Bajaj Allianz Life Goal Based Saving III",
                "Bajaj Allianz Life Elite Assure"
            ],
            "ulip_pension": [
                "Bajaj Allianz Life Smart Pension"
            ],
            "endowment_plans": [
                "Bajaj Allianz Life Assured Wealth Goal Platinum",
                "Bajaj Allianz Life ACE",
                "Bajaj Allianz Life Goal Suraksha"
            ]
        }
        
        return product_alignment
        
    except Exception as e:
        print(f"Error loading product alignment: {str(e)}")
        return {}
