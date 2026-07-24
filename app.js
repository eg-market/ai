let htmlScanner = null;

const statusBox = document.getElementById("status");
const productBox = document.getElementById("product");


// =========================
// Camera
// =========================

document
.getElementById("startCamera")
.addEventListener("click", startCamera);

document
.getElementById("stopCamera")
.addEventListener("click", stopCamera);

async function startCamera(){

    productBox.innerHTML="";
    statusBox.innerHTML="Starting camera...";

    if(htmlScanner){
        try{
            await htmlScanner.stop();
        }catch(e){}
    }

    htmlScanner = new Html5Qrcode("reader");

    try{

        await htmlScanner.start(

            {
                facingMode:"environment"
            },

            {
                fps:10,
                qrbox:{
                    width:300,
                    height:150
                }
            },

            onBarcode,

            ()=>{}

        );

        statusBox.innerHTML="Camera Ready";

    }
    catch(err){

        statusBox.innerHTML=err;

    }

}


async function stopCamera(){

    if(htmlScanner){

        try{

            await htmlScanner.stop();

            await htmlScanner.clear();

        }
        catch(e){}

    }

    document.getElementById("reader").innerHTML="";

    statusBox.innerHTML="";

}



// =========================
// Scan Image
// =========================

document
.getElementById("imageFile")
.addEventListener("change", scanImage);


async function scanImage(e){

    productBox.innerHTML="";
    statusBox.innerHTML="Reading image...";

    const file=e.target.files[0];

    if(!file)
        return;

    const reader=new ZXing.BrowserMultiFormatReader();

    const img=new Image();

    img.onload=async()=>{

        try{

            const result=await reader.decodeFromImageElement(img);

            onBarcode(result.text);

        }
        catch(err){

            console.log(err);

            statusBox.innerHTML="No barcode detected.";

        }

    };

    img.src=URL.createObjectURL(file);

}



// =========================
// Barcode detected
// =========================

async function onBarcode(code){

    statusBox.innerHTML="Barcode : "+code;

    if(htmlScanner){

        try{

            await htmlScanner.stop();

            await htmlScanner.clear();

        }catch(e){}

    }

    loadProduct(code);

}



// =========================
// OpenFoodFacts
// =========================

async function loadProduct(barcode){

    productBox.innerHTML="Loading...";

    const r=await fetch("/api/barcode/"+barcode);

    const p=await r.json();

    if(!p.success){

        productBox.innerHTML="<h2>"+p.message+"</h2>";

        return;

    }

    let h="";

    h+="<div class='card'>";

    if(p.image){

        h+="<img src='"+p.image+"'>";

    }

    h+="<h2>"+(p.name||"")+"</h2>";

    h+="<table>";

    add("Barcode",barcode);

    add("Arabic",p.name_ar);

    add("Brand",p.brand);

    add("Manufacturer",p.manufacturer);

    add("Categories",p.categories);

    add("Quantity",p.quantity);

    add("Countries",p.countries);

    add("Labels",p.labels);

    add("Ingredients",p.ingredients);

    add("Allergens",p.allergens);

    add("NutriScore",p.nutriscore);

    add("Nova",p.nova);

    add("EcoScore",p.ecoscore);

    h+="</table>";

    h+="<h3>Nutrition</h3>";

    h+="<pre>";

    h+=JSON.stringify(p.nutriments,null,2);

    h+="</pre>";

    h+="</div>";

    productBox.innerHTML=h;


    function add(a,b){

        if(!b)
            return;

        h+="<tr>";

        h+="<td><b>"+a+"</b></td>";

        h+="<td>"+b+"</td>";

        h+="</tr>";

    }

}