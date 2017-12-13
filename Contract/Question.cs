using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contract
{
    public class Question : INotifyPropertyChanged
    {
        public Question()
        {

        }
        private string _imagesource;
        public string ImageSource
        {
            get
            {
                return _imagesource;
            }
            set
            {
                this._imagesource = value;
                DoPropertyChanged("ImageSource");
            }
        }
        private int size;
        public int Size
        {
            get { return size; }
            set
            {
                this.size = value;
                DoPropertyChanged("Size");
            }
        }

        const string COLORTRUE = "#FF2E6C47";
        const string COLORFALSE = "#FFE84B43";
        const string HIDDDEN = "Hidden";
        const string VISIBILITY = "Visible";

        private string question;
        public string QuestionText
        {
            get
            {
                return question;
            }
            set
            {
                this.question = value;
                DoPropertyChanged("Question");
            }
        }
        public List<Option> options { get; set; }
        public Question(string question, List<Option> listOptions, string imagesource, int size)
        {
            this.question = question;
            this.options = listOptions;
            ImageSource = imagesource;
            this.size = size;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void DoPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public bool CheckAnswer(int i)
        {
            setColorAndVisible();
            if (options[i].Win == true)
            {
                return true;
            }
            return false;
        }
        public void setColorAndVisible()
        {
            for (int i = 0; i < options.Count; i++)
            {
                if (options[i].Win == true)
                {
                    options[i].Color = COLORTRUE;
                    if (options[i].Explanation != null)
                    {
                        options[i].Visible = VISIBILITY;
                    }
                }
                else
                {
                    options[i].Color = COLORFALSE;
                }

            }

        }


    }
}
