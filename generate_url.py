import codes
import creds

def generateURL():
    CODE_VERFIER, CODE_CHALLENGE = codes.generate_codes()
    OAUTH_URL = f'https://www.fitbit.com/oauth2/authorize?client_id={creds.CLIENT_ID}&response_type=code&code_challenge={CODE_CHALLENGE}&code_challenge_method=S256&scope=activity%20heartrate%20location%20nutrition%20oxygen_saturation%20profile%20respiratory_rate%20settings%20sleep%20social%20temperature%20weight%20electrocardiogram%20cardio_fitness%20social'
    print(CODE_VERFIER)
    print(OAUTH_URL)

generateURL()