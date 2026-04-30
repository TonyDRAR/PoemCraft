import os
import time
from urllib import error, parse, request


class PollinationsService:
    def __init__(self, api_url: str = "https://image.pollinations.ai"):
        self.api_url = api_url.rstrip("/")

    def generate_image(self, prompt: str, destination_path: str):
        encoded_prompt = parse.quote(prompt)
        query = parse.urlencode(
            {
                "width": 512,
                "height": 768,
                "nologo": "true",
                "model": "flux",
                "seed": time.time_ns() % 1000000000,
            }
        )
        url = (
            f"{self.api_url}/prompt/{encoded_prompt}"
            f"?{query}"
        )

        http_request = request.Request(
            url,
            headers={"User-Agent": "PoetryEditor/1.0"},
            method="GET",
        )

        try:
            with request.urlopen(http_request, timeout=90) as response:
                image_data = response.read()
        except error.HTTPError as exc:
            if exc.code == 429:
                raise RuntimeError(
                    "Pollinations limite temporairement les generations. Attendez un peu avant de reessayer."
                ) from exc

            raise RuntimeError(f"Pollinations a refuse la requete HTTP {exc.code}.") from exc
        except error.URLError as exc:
            raise RuntimeError(
                "Impossible de joindre Pollinations. Verifiez votre connexion internet."
            ) from exc
        except TimeoutError as exc:
            raise RuntimeError("Pollinations met trop de temps a repondre.") from exc

        if not image_data:
            raise RuntimeError("Pollinations n'a renvoye aucune image.")

        os.makedirs(os.path.dirname(destination_path), exist_ok=True)

        try:
            with open(destination_path, "wb") as image_file:
                image_file.write(image_data)
        except OSError as exc:
            raise RuntimeError(f"Impossible d'enregistrer l'image generee: {exc}") from exc
