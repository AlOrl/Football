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
        string lowResult = "Вы знаете много неожиданных фактов о домашних питомцах. Но сколько интересного еще предстоит узнать!\n"+
"PURINA PRO PLAN поможет вам прокачиваться в том же духе! :)";
        string lowState = "свещённый";
        string goodResult = "В вашем смартфоне всегда открыты пару вкладок с новостями из мира кошек и собак. Ведь продвинутый ветеринар всегда в курсе последних событий!\n"+
"PURINA PRO PLAN поддерживает вас, продолжайте прогрессировать каждый день! :)";
        string goodState = "двинутый";
        string excellentResult = "Вы просто гений, аплодируем стоя! Как настоящий эксперт, вы интересуетесь всеми аспектами своей профессии и можете рассказать массу интересного о пушистых любимцах.\n"+
"PURINA PRO PLAN в восторге от вашей эрудиции! :)";
        string excellentState = "грессивный";
        private string resultText;
        private string resultMessage;
        private Question currentquestion;
        private List<Question> questions;
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


        public TestViewModel()
        {
            newTest();
            ResultMessage = excellentResult;
            ResultText = excellentState;
        }

        private void newTest()
        {
            Index = 0;
            countTrueAnser = 0;
            LoadAsk t = new LoadAsk();
            questions = t.GetListQuestion();
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
                    ResultText = lowState;
                }
                if (CountTrueAnser >= 4 && CountTrueAnser < 8)
                {
                    ResultMessage = goodResult;
                    ResultText = goodState;
                }
                if (CountTrueAnser >= 8 )
                {
                    ResultMessage = excellentResult;
                    ResultText = excellentState;
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

        public Question CurrentQuestion
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
