import pkce

def generate_codes():
    code_verifier = pkce.generate_code_verifier(length=128)
    code_challenge = pkce.get_code_challenge(code_verifier)
    return [code_verifier,code_challenge]
