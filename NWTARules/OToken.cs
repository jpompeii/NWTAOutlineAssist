namespace NWTARules
{
    public class OToken
    {
        private string cellText;
        private int currentPos = 0;
        public bool LastToken { get; set; } = false;
        public string CurrToken { get; set; } = null;
        public int TokenCount { get; set; } = 0;
        public bool IsRule { get; set; } = false;
        public OToken(string _cellText)
        {
            if (String.IsNullOrEmpty(_cellText))
            {
                LastToken = true;
                CurrToken = String.Empty;
            }
            else
            {
                cellText = _cellText;
            }
        }

        public string lex()
        {
            if (LastToken)
                return null;

            IsRule = false;
            int i = cellText.IndexOf('{', currentPos);
            if (i > currentPos)
            {
                // found text
                CurrToken = cellText.Substring(currentPos, i - currentPos);
                TokenCount++;
                currentPos = i;
                char nextChar = cellText[i + 1];
                if (((i + 1) < cellText.Length) &&  nextChar == '{')
                {
                    CurrToken += "{";
                    currentPos += 2;
                }
                
            }
            else if (i == currentPos)
            {
                // found rule
                i = cellText.IndexOf('}', currentPos);
                if (i > currentPos)
                {
                    CurrToken = cellText.Substring(currentPos + 1, i - currentPos - 1);
                    TokenCount++;
                    currentPos = i + 1;
                    IsRule = true;
                }
            }
            else
            {
                // no more rules
                CurrToken = cellText.Substring(currentPos);
                TokenCount++;
                currentPos = cellText.Length;
            }
            if (currentPos >= cellText.Length)
                LastToken = true;

            return CurrToken;
        }

    }
}