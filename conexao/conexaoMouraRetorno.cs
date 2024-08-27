using System;
using System.Data;
using System.Data.SqlClient;

namespace CRM.conexao
{
    internal class conexaoMouraRetorno
    {
        private readonly string connectionString = "Server=cloudecia.jnmoura.com.br,1504;Database=CAESECIAMG_10221;User Id=CAESECIAMG_10221_POWERBI;Password=KI8msYRpRsRJifEw2ouw;TrustServerCertificate=True;";

        public DataTable FetchData(List<string> codigos, DateTime startDate, DateTime endDate)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    string query = @"
                        SELECT 
                            p.codigo,
                            p.nome,
                            p.fone,
                            p.fone2,
                            vi.Produto,
                            pr.Nome AS Nome_Produto,
                            CONVERT(varchar, v.data, 103) AS [Data da Venda],
                            vi.quantidade AS [Quantidade do Item],
                            CAST(vi.valor_total AS Money) AS [Valor Total do Item],
                            e.fantasia AS Empresa,
                            u.nome AS Vendedor
                        FROM 
                            pessoa p
                        INNER JOIN 
                            cliente c ON p.Codigo = c.Pessoa
                        INNER JOIN 
                            venda v ON v.cliente = c.Pessoa
                        INNER JOIN 
                            venda_item vi ON vi.venda = v.Codigo
                        INNER JOIN 
                            produto pr ON vi.Produto = pr.codigo
                        INNER JOIN 
                            empresa e ON e.codigo = v.empresa
                        INNER JOIN 
                            usuario u ON v.vendedor = u.codigo
                        WHERE 
                            v.data BETWEEN @startDate AND @endDate
                            AND p.codigo IN (" + string.Join(",", codigos.Select((c, i) => $"@codigo{i}")) + @")
                        ORDER BY 
                            vi.venda DESC;
                    ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@startDate", startDate.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@endDate", endDate.ToString("yyyy-MM-dd"));

                        for (int i = 0; i < codigos.Count; i++)
                        {
                            cmd.Parameters.AddWithValue($"@codigo{i}", codigos[i]);
                        }

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao buscar os dados: " + ex.Message);
            }

            return dataTable;
        }
    }
}