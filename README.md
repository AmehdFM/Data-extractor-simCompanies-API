# SimCompanies 
scrip que recopila información de algunos recursos del juego simCompanies mediante su API y los guarda en un archivo .xlsx

## analizar_mercado.py

Analizador de mercado que usa la API pública de [Simco Tools](https://api.simcotools.com)
para decidir qué recursos conviene **comprar**, **vender** o **vigilar**.

Para cada combinación (recurso, calidad):

1. Descarga el histórico de velas diarias y calcula el promedio (VWAP ponderado
   por volumen), la media móvil de 7 días y la tendencia por regresión lineal.
2. Consulta el precio actual.
3. Comprueba la rentabilidad: `0.32 * transportation + compra <= venta * (1 - 4%)`.
4. Clasifica el ítem en `COMPRAR AHORA`, `BAJISTA - VIGILAR` o `ALCISTA - VENDER`
   y lo muestra en una tabla ordenada.

```bash
pip install requests pandas numpy scipy tabulate
python analizar_mercado.py
python analizar_mercado.py --resources 74,1,10 --qualities 0,1,2 --csv salida.csv --quiet
```

Los parámetros (realm, lista de recursos, calidades, umbrales, impuesto, coste de
transporte) están en la sección `CONFIGURACIÓN` al inicio del script.
