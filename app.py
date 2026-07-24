from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import requests

app = FastAPI()   # <-- أولاً

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")   # <-- بعد إنشاء app


OFF_URL = "https://world.openfoodfacts.org/api/v2/product/{}.json"


@app.get("/barcode/{barcode}")
def barcode(barcode: str):

    r = requests.get(
        OFF_URL.format(barcode),
        headers={"User-Agent": "EgyptMarket/1.0"},
        timeout=20,
    )

    if r.status_code != 200:
        raise HTTPException(500, "OpenFoodFacts unavailable")

    j = r.json()

    if j.get("status") != 1:
        return {
            "found": False,
            "barcode": barcode
        }

    p = j["product"]

    return {
        "found": True,
        "barcode": barcode,

        "name": p.get("product_name"),
        "name_ar": p.get("product_name_ar"),
        "generic_name": p.get("generic_name"),

        "brand": p.get("brands"),
        "manufacturer": p.get("manufacturing_places"),
        "countries": p.get("countries"),

        "categories": p.get("categories"),
        "labels": p.get("labels"),
        "quantity": p.get("quantity"),
        "packaging": p.get("packaging"),

        "ingredients": p.get("ingredients_text"),

        "allergens": p.get("allergens"),
        "traces": p.get("traces"),

        "nova_group": p.get("nova_group"),
        "nutriscore": p.get("nutriscore_grade"),
        "ecoscore": p.get("ecoscore_grade"),

        "image": p.get("image_front_url"),
        "image_front": p.get("image_front_url"),
        "image_ingredients": p.get("image_ingredients_url"),
        "image_nutrition": p.get("image_nutrition_url"),
        "image_packaging": p.get("image_packaging_url"),

        "nutriments": p.get("nutriments", {})
    }