using System.Data;
using System.Globalization;
using System.Windows.Forms;
using Npgsql;

namespace Starline
{
    class ConexaoNpgSql
    {
        public NpgsqlConnection Conn;
        private NpgsqlTransaction Transacao;
        public ConexaoNpgSql()
        {
            try
            {
                string DadosConexao = System.IO.File.ReadAllText(@"C:\Arquivos\DadosConexao.str");
                string connectiosStrings = ProcessTest.Decriptar(DadosConexao);
                Conn = new NpgsqlConnection(connectiosStrings);


            }
            catch
            {

            }
        }

        public void AbrirConexao()
        {
            if (!ConexaoAberta())
            {
                try
                {
                    Conn.Open();
                }
                catch
                {

                }
            }
        }
        public void FecharConexao()
        {
            if (ConexaoAberta())
            {
                Conn.Close();
            }
        }

        public bool ConexaoAberta()
        {
            return Conn.State == System.Data.ConnectionState.Open;

        }

        public NpgsqlDataReader ExecutarSelect(string select)
        {
            if (!ConexaoAberta())
            {
                AbrirConexao();
            }
            NpgsqlCommand cmd = new NpgsqlCommand(select, Conn);
            return cmd.ExecuteReader();
        }

        public void IniciaTransacao()
        {

            if (!ConexaoAberta())
            {
                AbrirConexao();
            }
            Transacao = Conn.BeginTransaction(IsolationLevel.Serializable);
        }

        public bool TransacaoAtiva()
        {
            return Transacao != null;
        }

        public void Commit()
        {
            if (TransacaoAtiva())
            {
                Transacao.Commit();
            }
        }
        public void Rollback()
        {
            if (TransacaoAtiva())
            {
                Transacao.Rollback();
            }
        }
        public void ExecutarScript(string script)
        {
            if (!ConexaoAberta())
            {
                AbrirConexao();
            }
            try
            {
                NpgsqlCommand cmd = new NpgsqlCommand(script, Conn);
                cmd.ExecuteNonQuery();

            }
            finally
            {
                if (ConexaoAberta())
                {
                    FecharConexao();
                }
            }
        }

        public void GetDadosTodos(string script, DataGridView dgv)
        {
            if (!ConexaoAberta())
            {
                AbrirConexao();
            }
            try
            {
                NpgsqlDataAdapter dataAdapter = new NpgsqlDataAdapter(script, Conn);
                NpgsqlCommandBuilder buider = new NpgsqlCommandBuilder(dataAdapter);
                DataTable dtable = new DataTable()
                {
                    Locale = CultureInfo.InvariantCulture
                };
                dataAdapter.Fill(dtable);
                dgv.DataSource = dtable;
            }
            finally
            {
                if (ConexaoAberta())
                {
                    FecharConexao();
                }
            }
        }

    }
}
