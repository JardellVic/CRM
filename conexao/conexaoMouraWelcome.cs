using System;
using System.Data;
using System.Data.SqlClient;

namespace CRM.conexao
{
    public class conexaoMouraWelcome
    {
        // Definindo a string de conexão dentro da classe
        private readonly string connectionString = "Server=cloudecia.jnmoura.com.br,1504;Database=CAESECIAMG_10221;User Id=CAESECIAMG_10221_POWERBI;Password=KI8msYRpRsRJifEw2ouw;TrustServerCertificate=True;";

        public DataTable FetchData(DateTime startDate, DateTime endDate)
        {
            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"
                                SELECT C.codigo AS CodCli,
                           C.*,
                           CI.cidade AS Nome_Cidade,
                           C.obs,
                           C.contato_cliente,
                           CI.estado,
                           C.classe,
                           E.fantasia AS Empresa,
                           U.abreviacao AS Nome_Vendedor,
                           S.descricao AS Segmento,
                           Z.descricao AS Zona_Venda,
                           C.observacao_endereco,
                           (SELECT Count(codigo) AS Qtde_Vendas
                            FROM venda
                            WHERE cliente = C.codigo
                                  AND cancelada = 'N') quantidade,
                           (SELECT TOP 1 data
                            FROM venda
                            WHERE cliente = C.codigo
                            ORDER BY data DESC) AS Data_Ultima_Venda
                    FROM pessoa C
                    INNER JOIN cliente Cli ON C.codigo = Cli.pessoa
                    LEFT JOIN empresa E ON E.codigo = C.empresa_origem
                    LEFT JOIN cidade CI ON CI.codigo = C.cidade
                    LEFT JOIN usuario U ON U.codigo = C.vendedor
                    LEFT JOIN segmento S ON S.codigo = C.segmento
                    LEFT JOIN zona_venda Z ON Z.codigo = C.zona_venda
                    LEFT JOIN representante RE ON C.representante = RE.codigo
                    LEFT JOIN convenio CO ON CO.codigo_cliente = C.codigo
                    WHERE CO.codigo_cliente IS NULL
                          AND C.Data_Cadastro BETWEEN @startdate AND @enddate
                    ORDER BY C.nome;

                ";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@startDate", startDate);
                cmd.Parameters.AddWithValue("@endDate", endDate);

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }
            return dt;
        }
    }
}
