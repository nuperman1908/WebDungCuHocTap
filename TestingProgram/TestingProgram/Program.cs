namespace TestingProgram;

static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new FormAddProduct_Test());
    }
    
}