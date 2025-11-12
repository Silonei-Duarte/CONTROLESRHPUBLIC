import oracledb

# Configuração do banco de dados
DB_CONFIG = {
    'host': '*',
    'port': '*',
    'service_name': '*',
    'user': '*',
    'password': '*'
}

# Consulta SQL a ser testada
SQL_QUERY = """
SELECT 
    F.NUMEMP, 
    H.CODLOC, 
    F.NUMCAD, 
    F.NOMFUN,
    N.DESNAC,
    TO_CHAR(F.DVLEST, 'DD/MM/YYYY') AS DVLEST,
    TO_CHAR(E.DATTER, 'DD/MM/YYYY') AS DATTER,
    E.VISEST
FROM R034FUN F
JOIN R016HIE H 
    ON F.NUMLOC = H.NUMLOC
JOIN R023NAC N 
    ON F.CODNAC = N.CODNAC
JOIN R034EST E 
    ON F.NUMEMP = E.NUMEMP
   AND F.TIPCOL = E.TIPCOL
   AND F.NUMCAD = E.NUMCAD
WHERE F.NUMEMP IN (10, 16, 17, 18, 19)
  AND F.SITAFA NOT IN ('007')
  AND F.CODNAC NOT IN (10)

"""


def formatar_resultados_sem_tabulate(results, headers):
    """Formata e exibe os resultados manualmente, sem tabulate."""
    # Imprime os cabeçalhos
    print(" | ".join(headers))
    print("-" * (len(" | ".join(headers)) + 4))

    # Imprime cada linha do resultado
    for row in results:
        formatted_row = [
            f"{col:.10f}" if isinstance(col, float) else str(col) for col in row
        ]
        print(" | ".join(formatted_row))


def main():
    try:
        # Conectando ao banco
        dsn = oracledb.makedsn(
            DB_CONFIG['host'],
            DB_CONFIG['port'],
            service_name=DB_CONFIG['service_name']
        )
        connection = oracledb.connect(
            user=DB_CONFIG['user'],
            password=DB_CONFIG['password'],
            dsn=dsn
        )
        print("Conexão estabelecida com sucesso.")

        # Executando a consulta
        with connection.cursor() as cursor:
            cursor.execute(SQL_QUERY)

            if cursor.description:  # é SELECT
                column_names = [desc[0] for desc in cursor.description]
                results = cursor.fetchall()
                if results:
                    formatar_resultados_sem_tabulate(results, column_names)
                else:
                    print("Nenhum resultado encontrado.")
            else:  # é UPDATE/INSERT/DELETE
                connection.commit()
                print(f"{cursor.rowcount} linha(s) afetada(s).")

    except oracledb.DatabaseError as e:
        print(f"Erro ao conectar ou executar a consulta: {e}")
    finally:
        if 'connection' in locals() and connection:
            connection.close()
            print("Conexão fechada.")


if __name__ == "__main__":
    main()
