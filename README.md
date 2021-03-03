
## **Cronograma de Pagos**

Las celdas en Color Verde son las celdas que se modifican, el resto no se modifican en casos normales.
En caso Gracia Total, los gastos se duplican hasta cancelar lo pendiente, modificacion manual solo necesario si no tiene mas cuotas pendientes de pago.
Automaticamente se cambia el relleno en Naranja en caso se encuentre una advertencia y en Rojo en caso Valor Negativo en celdas importantes

***Columna A - Tipo Cuota:***

 - Regular: Paga Intereses + Gastos y Amortiza el resto
 - Gracia Parcial:  Paga Intereses + Gastos
 - Prepago: Mismo comportamiento que regular pero no corre la cantidad de cuotas

***Columna B:*** Numero de Cuotas Gracia Total

***Columna C:*** Fijar Cuota

***Columna D:*** Fijar Intereses
Nota: No afecta intereses acumulados

***Columna E - Fijar Fecha***

***Rango L1:Y2:***
Para definir cuotas dobles por mes

***Celda N6:*** Fecha desembolso

***Celda N7:*** Dia de pago

***Celda N8:*** Monto desembolsado

***CheckBox Mes adicional:***
Para correr un mes adicional en la primera cuota, check en caso de que la primera cuota sea cercana al desembolso

***Celda N8:***
Monto desembolsado

***Celda Q6:***
Numero de cuotas Original

***Celda Q7:***
Numero de cuotas Adicionales

Se usa en caso se halla otorgado periodos de gracia post-desembolso o prepagos

***Boton Ajustar***

Presionar este boton tras modificar todo los parametros deseados
Tiene vinculada una macro que intenta ajustar la cuota a fuerza bruta.

***Boton Exportar .csv***

Genera un .csv con el formato aceptado por SISGO.
Lo muestra en pantalla, lo guarda en la carpeta "C:\Cronogramas\" y con el nombre de la fecha y hora actual (YYYYMMDD-HHmmss).
