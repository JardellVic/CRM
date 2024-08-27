using System;
using System.Data;
using System.Data.SqlClient;

namespace CRM.conexao
{
    public class conexaoMouraVacina
    {
        // Definindo a string de conexão dentro da classe
        private readonly string connectionString = "Server=cloudecia.jnmoura.com.br,1504;Database=CAESECIAMG_10221;User Id=CAESECIAMG_10221_POWERBI;Password=KI8msYRpRsRJifEw2ouw;TrustServerCertificate=True;";

        public DataTable FetchData(DateTime startDate, DateTime endDate)
        {
            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"
                    SELECT DISTINCT
                        M.codigo,
                        M.servico AS Cod_Servico,
                        M.data,
                        M.dose,
                        M.atendimento,
                        A.nome_animal,
                        M.animal AS Cod_Animal,
                        C.codigo AS Cod_Proprietario,
                        C.nome AS Proprietario,
                        C.email AS E_mail_Proprietario,
                        E.razao_social AS Empresa,
                        C.fone,
                        S.nome AS Servico,
                        M.proprietario AS Cod_Proprietario,
                        COALESCE(A.inativo, 'N') AS Inativo
                    FROM 
                        animais_marketing_servicos M
                    INNER JOIN 
                        animais_cadastro A ON A.codigo = M.animal
                    INNER JOIN 
                        pessoa C ON C.codigo = M.proprietario
                    INNER JOIN 
                        servico S ON S.codigo = M.servico
                    LEFT JOIN 
                        empresa E ON E.codigo = M.empresa
                    WHERE 
                        M.data >= @startdate AND M.data <= @enddate
                    ORDER BY 
                        M.data ASC;
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
