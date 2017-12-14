using BL;
using Contract;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestForGods
{

    public class TestViewModel : INotifyPropertyChanged
    {
        string lowResult = "Похоже, Вы не очень то интересуетесь футболом. Но ничего, до Чемпионата Мира еще достаточно времени, чтобы наверстать упущенное:)";
        string goodResult = "Весьма неплохо! Еще немного, и Вас можно будет считать гуру футбольной истории!";
        string excellentResult = "Вы настоящий футбольный эксперт, спортивным журналистом не хотите подработать?";
        private string resultText;
        private string resultMessage;
        private Contract.Question currentquestion;
        private List<Contract.Question> questions;
        private int index;
        private int countTrueAnser;
        private bool step = false;
        private string finalPageVisibility = "Collapsed";
        private string beginVisibility = "Visible";
        private string thickness = "3";
        public int Index
        {
            get
            {
                return index;
            }
            set
            {
                this.index = value;
                DoPropertyChanged("Index");
            }
        }
        public string ResultText
        {
            get { return resultText; }
            set { resultText = value;
                DoPropertyChanged("ResultText");
            }
        }
        public string ResultMessage
        {
            get { return resultMessage; }
            set
            {
                resultMessage = value;
                DoPropertyChanged("ResultMessage");
            }
        }
        public string FinalPageVisibility
        {
            get { return finalPageVisibility; }
            set
            {
                finalPageVisibility = value;
                DoPropertyChanged("FinalPageVisibility");
            }
        }
        public string BeginVisibility
        {
            get { return beginVisibility; }
            set
            {
                beginVisibility = value;
                DoPropertyChanged("BeginVisibility");
            }
        }
        public string Thickness
        {
            get { return thickness; }
            set
            {
                thickness = value;
                DoPropertyChanged("Thickness");
            }
        }
        public int CountTrueAnser
        {
            get
            {
                return countTrueAnser;
            }
            set
            {
                this.countTrueAnser = value;
                DoPropertyChanged("CountTrueAnser");
            }
        }

        //Исправлено 
        public TestViewModel()
        {
            provider = new DataProcessing();
            newTest();
            ResultMessage = excellentResult;
           
        }
        private DataProvider provider;


        private void newTest()
        {
            Index = 0;
            countTrueAnser = 0;
            questions =  provider.GetQuestions();
            move();
            step = false;
   
        }

        public void move()
        {
            Thickness = "3";
            Index++;
            if (index < 11)
            {
                step = false;
                CurrentQuestion = questions[index - 1];
            }
            else
            {
                if(CountTrueAnser < 4)
                {
                    ResultMessage = lowResult;

                }
                if (CountTrueAnser >= 4 && CountTrueAnser < 8)
                {
                    ResultMessage = goodResult;

                }
                if (CountTrueAnser >= 8 )
                {
                    ResultMessage = excellentResult;

                }
                FinalPageVisibility = "Visible";
                //открытие результатов 
            }


        }

        public void choose(int i)
        {
            Thickness = "0";
            if (step != true)
                if (currentquestion.CheckAnswer(i))
                {
                    CountTrueAnser++;
                    step = true;
                }
        }

        public Contract.Question CurrentQuestion
        {
            get
            {
                return currentquestion;
            }
            set
            {
                this.currentquestion = value;
                DoPropertyChanged("CurrentQuestion");
            }
        }
        //переход к начальной странице
        private RelayCommand newGame;
        public RelayCommand NewGame
        {
            get
            {
                return newGame ??
                  (newGame = new RelayCommand(obj =>
                  {
                      newTest();
                      FinalPageVisibility = "Collapsed";
                      BeginVisibility = "Visible";
                  }));
            }
        }
        //переход
        private RelayCommand next;
        public RelayCommand Next
        {
            get
            {
                return next ??
                  (next = new RelayCommand(obj =>
                  {
                      move();
                  }));
            }
        }

        // выбор ответа 
        private RelayCommand click;
        public RelayCommand Click
        {
            get
            {
                return click ??
                  (click = new RelayCommand(obj =>
                  {
                      int ans = int.Parse(obj.ToString());
                      choose(ans);

                  }));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void DoPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
