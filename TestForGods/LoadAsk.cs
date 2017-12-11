using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestForGods
{
   public class LoadAsk
    {
        private List<Question> ListOfAsk { get; set; }
        public LoadAsk()
        {
            ListOfAsk = new List<TestForGods.Question>()
          {
              new Question ("Какая страна выигрывала Чемпионат Мира наибольшее количество раз?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Бразилия","Бразильская сборная становилась триумфатором аж 5 раз: в 1958, 1962, 1970, 1994 и 2002 годах", true),
                new Option("Германия", null, false),
                new Option ("Италия", null, false),
                new Option ("Испания", null, false),
              },
              "Images/ImageForTest/1.jpg",
              54
              ),
              new Question ("Кто является лучшим бомбардиром в истории ЧМ?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Роналдо","null", false),
                new Option("Герд Мюллер", null, false),
                new Option ("Пеле", null, false),
                new Option ("Мирослав Клозе","Ветеран немецкой сборной наколотил 16 мячей", true),
              },   "Images/ImageForTest/2.jpg",55),
              new Question ("Сборная какой страны трижды выходила в финал ЧМ, но так ни разу и не завладела трофеем?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Уругвай",null, false),
                new Option("Португалия", null, false),
                new Option ("Нидерланды", "Оранжевые играли в финалах в 1978, 1982 и 2010 годах, но все три попытки оказались неудачными", true),
                new Option ("Англия", null, false),
              },  "Images/ImageForTest/3.jpg", 54),
              new Question ("Сборная какой страны победила в первом в истории ЧМ?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Бразилия",null, false),
                new Option("Германия", null, false),
                new Option ("Уругвай", "На домашнем ЧМ в финале они одолели Аргентину 4:2 ", true),
                new Option ("Италия", null, false),
              },   "Images/ImageForTest/4.jpg", 55),
              new Question ("Только одна сборная в истории, будучи хозяйкой Чемпионата, не смогла выйти из группы. Какая?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Чили",null, false),
                new Option("Япония", null, false),
                new Option ("ЮАР", "В 2010 году на домашнем ЧМ они заняли последнюю строчку в группе, проиграв все три матча", true),
                new Option ("Швеция", null, false),
              },   "Images/ImageForTest/5.jpg",54),
              new Question ("Гол, названный впоследствии рукой бога забил...", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Пеле",null, false),
                new Option ("Ференц Пушкаш", null, false),
                new Option("Диего Марадона", "Великий и неповторимый:)", true),
                new Option ("Роналдиньо", null, false),
              },   "Images/ImageForTest/6.jpg", 54),
              new Question ("Только один российский футболист становился лучшим бомбардиром ЧМ. Кто?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Никита Симонян",null, false),
                new Option ("Олег Саленко", "В 1994 году он забил 6 мячей, но даже это, увы, не помогло Российской сбоной выйти из группы", true),
                new Option("Александр Кержаков", null, false),
                new Option ("Александр Мостовой", null, false),
              },   "Images/ImageForTest/7.jpg", 55),
              new Question ("Сборная СССР всего раз в истории добаралась до полуфинала ЧМ-в 1966 году. Кто тогда остановил их чемпионскую поступь?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Англия",null, false),
                new Option ("Бразилия",null, false),
                new Option("Аргентина", null, false),
                new Option ("ФРГ","2:1 закончилась та встреча", true),
              },   "Images/ImageForTest/8.jpg", 42)
              ,
              new Question ("Вратарь сбороной Чехословакии Франтишек Планичка в матче с сборной Бразилии в 1938 году 35 минут играл ?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("с сотрясением мозга",null, false),
                new Option ("без перчаток",null, false),
                new Option("со сломанной рукой", "В это сложно поверить, но Планичка действительно доигрывал с сломанной рукой,и при этом не пропустил ни одного мяча", true),
                new Option ("в нападении", null, false),
              },   "Images/ImageForTest/9.jpg", 46)
              ,
              new Question ("Финальный вопрос: лучший бомбардир сборной России по футболу в истории?", new List<TestForGods.Option>()
              {
                new TestForGods.Option("Александр Мостовой",null, false),
                new Option ("Александр Кержаков","Несмотря на все промахи, Александр успел наколотить в составе сборной 30 мячей", true),
                new Option("Олег Саленко", null, false),
                new Option ("Александр Кокорин", null, false),
              },   "Images/ImageForTest/10.jpg", 42)
          };
        }
        public List<Question> GetListQuestion()
        {
            return ListOfAsk; 
        }
    }
}
