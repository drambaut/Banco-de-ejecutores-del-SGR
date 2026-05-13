# Decisiones metodológicas

## Feature Selection — Pipeline No Supervisado

### Criterios de eliminación de variables (en orden de aplicación)
1. **Nulidad** (umbral 60%): columnas con más del 60% de valores nulos se descartan.
2. **Varianza cercana a cero** (umbral 0.01): columnas sin variabilidad útil.
3. **Multicolinealidad — VIF** (umbral 10): elimina variables redundantes.
4. **Laplacian Score** (k=5 vecinos): rankea variables por su capacidad de preservar estructura local.

### Por qué no supervisado
No se dispone de una variable objetivo (`y`) definida. El pipeline identifica variables
que capturan la mayor variabilidad y estructura de los datos sin etiquetas.

### Notebook principal
`notebooks/feature_selection/03_resultado_final.ipynb`
