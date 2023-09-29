import models
import Emails
if __name__ == "__main__":
    Emails.leer_email()                                 # 1. Lee los emails y descarga los adjuntos de servipag.
    models.eliminarcarpetas()                           # 2. Elimina carpetas para iniciar proceso.
    models.creacionCarpetas()                           # 3. Creacion de carpeta para iniciar el proceso.
    models.leerCarpetaPl()                              # 4. Leyendo todos los archivos que comience con PL.
    models.leerCarpetaHE()                              # 5. Leyendo todos los archivos que comience con HE.
    models.PL()                                         # 6. Inicio de ETL ,extrayendo los formato A4.
    models.HE()                                         # 7. Inicio de ETL ,extrayendo los formato los txt para servipag.
    models.UnionPLyHE()                                 # 8. Union de los datos del PL y HE.
    models.ETLA4_TRABAJADOR()                           # 9. Consulta para trabajadores de productos financieros ,creditos hipotecario y seguros.
    models.prelacionRevicion()                          # 10. Se aplica los formatos a los datos prelando la tablas.
    models.ETLA4_PENSIONADO()                           # 11. Consulta para pensionados de productos financieros ,creditos hipotecario y seguros.
    models.prelacionRevicion_PENSIONADO()               # 12. Se aplica los formatos a los datos prelando la tablas.
    models.creaciontxt_PLHR()                           # 13. generando la Salida del  PLHR RAM.
