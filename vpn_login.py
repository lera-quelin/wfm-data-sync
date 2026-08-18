import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =====================================================================
# 1. FONCTIONS UTILITAIRES (Issues de votre script)
# =====================================================================
def attendre_et_cliquer(driver, xpath, timeout=15):
    try:
        element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        element.click()
        return True
    except Exception as e:
        print(f"Erreur lors du clic sur {xpath} : {e}")
        return False

def attendre_et_ecrire(driver, xpath, texte, timeout=15):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element.clear()
        element.send_keys(texte)
        return True
    except Exception as e:
        print(f"Erreur lors de l'écriture sur {xpath} : {e}")
        return False

# =====================================================================
# 2. SÉQUENCE DE CONNEXION
# =====================================================================
def login_concentrix(user_id, user_pwd):
    print("\n-> Initialisation du navigateur Edge (headless)...")
    options = Options()
    options.add_argument("--headless=old") # Retirez pour voir la page s'afficher lors des tests
    options.add_argument("--window-position=-2400,-2400")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--log-level=3")
    
    driver = webdriver.Edge(options=options)
    URL_LOGIN = "https://signin.concentrix.com/admin/default/login"

    try:
        print(f"1. Accès à l'URL : {URL_LOGIN}")
        driver.get(URL_LOGIN)
        time.sleep(3) # Laisse le temps aux scripts de la page de s'initialiser

        print("2. Saisie de l'identifiant...")
        # Sélecteurs élargis pour capter les différents types de formulaires (Okta, Azure, Ping)
        attendre_et_ecrire(driver, "//input[@type='text' or @type='email' or contains(@name, 'login') or contains(@id, 'username')]", user_id)
        attendre_et_cliquer(driver, "//button[@type='submit' or contains(., 'Next') or contains(., 'Suivant')]")
        time.sleep(3)

        print("3. Saisie du mot de passe...")
        attendre_et_ecrire(driver, "//input[@type='password']", user_pwd)
        attendre_et_cliquer(driver, "//button[@type='submit' or contains(., 'Sign In') or contains(., 'Connexion') or contains(., 'Verify')]")
        time.sleep(5)

        # OPTIONNEL : Gérer le MFA ou la popup "Rester connecté ?"
        print("4. Vérification post-connexion (MFA ou confirmations)...")
        if attendre_et_cliquer(driver, "//input[@type='button' or @type='submit'][(contains(@value, 'Yes') or contains(@value, 'Oui'))]", timeout=5):
            print(" -> Popup 'Rester connecté' validée.")
            time.sleep(3)

        print("✅ Connexion réussie ! La session est active dans le navigateur.")
        
        # Vous pouvez maintenant enchaîner avec vos navigations ou extractions
        # Ex: driver.get("URL_DE_VOTRE_OUTIL_INTERNE")

        return driver # Retourne l'instance du navigateur pour la suite du script

    except Exception as e:
        print(f"\n❌ [ERREUR CRITIQUE] Échec de la connexion : {e}")
        driver.quit()
        return None

# =====================================================================
# 3. TEST DE LA FONCTION
# =====================================================================
if __name__ == "__main__":
    USER_ID = "votre_email@concentrix.com"
    USER_PWD = "votre_mot_de_passe"
    
    navigateur_actif = login_concentrix(USER_ID, USER_PWD)
    
    if navigateur_actif:
        print("Fermeture du navigateur de test...")
        navigateur_actif.quit()
