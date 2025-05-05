from supabase import create_client, Client
import os
import dotenv

class banco_supabase:

    def __init__(self):
        pass

dotenv.load_dotenv()

supabase_url = os.getenv("SUPABASE_URL")
supabase_key = os.getenv("SUPABASE_KEY")

supabase: Client = create_client(supabase_url, supabase_key)

response = supabase.table("lista_materiais").select("*").execute()
print(response)
