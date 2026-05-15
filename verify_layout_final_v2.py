import asyncio
from playwright.async_api import async_playwright
import os
import re

async def generate_screenshot():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()

        with open('codigo.gs', 'r') as f:
            content = f.read()

        # Find the content of the function generating the HTML
        match = re.search(r"return '(<!DOCTYPE html>.*?)';", content, re.DOTALL)
        if not match:
            print("Could not find HTML template")
            return

        html_template = match.group(1)

        # Define the styles to match codigo.gs
        F    = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;'
        FB   = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;'
        UL   = 'border-bottom:1px solid #000;'
        SEC  = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;text-align:center;padding:10px 0;'

        # Perform replacements
        html_final = html_template.replace("'+FB+'", FB)
        html_final = html_final.replace("'+F+'", F)
        html_final = html_final.replace("'+UL+'", UL)
        html_final = html_final.replace("'+SEC+'", SEC)
        html_final = html_final.replace("'+urlImagen+'", "https://via.placeholder.com/150x100?text=LOGO")
        html_final = html_final.replace("'+fDia+'", "15")
        html_final = html_final.replace("'+fMes+'", "05")
        html_final = html_final.replace("'+fAnio+'", "2026")
        html_final = html_final.replace("'+chkN+'", "X")
        html_final = html_final.replace("'+chkR+'", "&nbsp;&nbsp;")
        html_final = html_final.replace("'+PT+'", "QUINCENAL")
        html_final = html_final.replace("'+(r.Nombre_Promotor||'')+'", "JUAN PEREZ")
        html_final = html_final.replace("'+(r.Nombre_Trabajador||'')+'", "JOSE LOPEZ GARCIA")
        html_final = html_final.replace("'+(r.Domicilio||'')+'", "AV. PRINCIPAL 123, CUERNAVACA, MORELOS")
        html_final = html_final.replace("'+(r.Tel_Casa||'')+'", "7771234567")
        html_final = html_final.replace("'+(r.Tel_Celular||'')+'", "7779876543")
        html_final = html_final.replace("'+(r.Secretaria||'')+'", "SECRETARIA DE SALUD")
        html_final = html_final.replace("'+(r.Cargo||'')+'", "ADMINISTRATIVO")
        html_final = html_final.replace("'+(r.Clave||'')+'", "12345")
        html_final = html_final.replace("'+(r.RFC||'')+'", "LOGJ800101XXX")
        html_final = html_final.replace("'+(r.Monto_Solicitado||'')+'", "10,000.00")
        html_final = html_final.replace("'+(r.Plazo||'')+'", "24")
        html_final = html_final.replace("'+(r.Descuento||'')+'", "500.00")
        html_final = html_final.replace("'+(r.Total_Pagar||'')+'", "12,000.00")

        # Set content directly
        await page.set_content(html_final)
        await page.screenshot(path='verification_v4.png', full_page=True)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(generate_screenshot())
