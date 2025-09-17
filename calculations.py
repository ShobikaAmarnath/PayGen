# calculations.py
from num2words import num2words
from config import EPF_RATE, PROF_TAX

def convert_to_inr_words(num):
    """Converts a number to Indian currency in words."""
    rupees = int(num)
    paise = round((num - rupees) * 100)
    if rupees > 0:
        words = f"{num2words(rupees, lang='en_IN')} rupees"
    else:
        words = ""
    if paise > 0:
        paise_words = num2words(paise, lang='en_IN')
        if words:
            words += f" and {paise_words} paise"
        else:
            words = f"{paise_words} paise"
    if not words:
        return "Zero Rupees Only"
    return words.title() + " Only"

def calculate_salary(employee_data):
    """Calculates all salary components for a single employee."""
    gross = float(employee_data["Gross_Salary"])
    income_tax = float(employee_data["Income_Tax"])

    basic = 0.50 * gross # Basic is 50% of Gross
    hra = 0.40 * basic # HRA is 40% of Basic or 20% of Gross
    special_allowance = 0.30 * gross # Special Allowance is 30% of Gross
    
    epf = round(basic * EPF_RATE, 2)
    total_ded = epf + income_tax + PROF_TAX
    net = gross - total_ded
    net_in_words = convert_to_inr_words(net)

    return {
        "basic": basic, "hra": hra, "special_allowance": special_allowance, "gross": gross, 
        "epf": epf, "income_tax": income_tax, "prof_tax": PROF_TAX, "total_ded": total_ded, 
        "net": net, "net_in_words": net_in_words
    }