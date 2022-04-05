
namespace MarLang
{
    public class QuizClass
    {
        public string Word {get; private set;}
        public string[] Descriptions {get; private set;}
        public int CorrectAnswer {get; private set;}

        public QuizClass(string word, int correctAnswer, params string[] descriptions)
        {
            Word = word;
            CorrectAnswer = correctAnswer;
            Descriptions = descriptions;
        }
    }

    public class IsDeletedClass
    {
        public long chatId {get; private set;}
        public bool isDeleted {get;set;}

        public IsDeletedClass(long chatId)
        {
            this.chatId = chatId;
            isDeleted = false;
        }
    }
}