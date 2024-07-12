namespace LegacyExchangeDNConverter
{
    public class DebugConsole
    {
        public static void WriteLine(string message, ConsoleColor color = ConsoleColor.White)
        {
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write("[");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("*");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write("] ");

            if (color == ConsoleColor.Yellow)
                Console.BackgroundColor = color;
            else
            {
                Console.ForegroundColor = color;
            }
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public static void Write(string message, bool continueLine = false, ConsoleColor color = ConsoleColor.White)
        {
            if (!continueLine)
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write("[");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("*");
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write("] ");
            }

            Console.ForegroundColor = color;
            Console.Write(message);
        }
    }
}
