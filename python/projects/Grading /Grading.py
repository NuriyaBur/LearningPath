# coding: utf-8

# Это калькулятор определения грейда на основе Балльно-факторного метода: должности присваивается определенное количество баллов по каждому фактору. Итоговая сумма баллов определяет грейд с учетом весов факторов.
# В этой модели 3 основных фактора: КВАЛИФИКАЦИЯ,  СЛОЖНОСТЬ и  ОТВЕТСТВЕННОСТЬ
# КВАЛИФИКАЦИЯ определяется 4 подфакторами: Специальные знания и умения, Навыки общения, Знание бизнеса и Опыт работы
# СЛОЖНОСТЬ  определяется 2 подфакторами: Сложность решаемых вопросов и Область решаемых вопросов
# ОТВЕТСТВЕННОСТЬ определяется 3 подфакторами: Самостоятельность в принятии решений, Ответственность за финансовые результаты и Ответственность за работу других
# У каждого подфактора есть от 3 до 6 критериев оценки, которым определяется количество баллов
# разработано для фарм компании в 2020 году на основе методички Алексея Реброва https://www.delfy.biz/books


# импортируем библиотеки. Предварительно нужно установить библиотеку rich через pip 
import pandas as pd
from rich.console import Console
from rich.table import Table
import openpyxl


#Создадим критерии для подфактора "Специальные знания и умения"
Special_knowledge_and_skills = pd.DataFrame({'Сriteria_name': ['Достаточно среднего или н/высшего образования,специальных знаний не требуется', 
                                                               'Необходимо высшее образование, не обязательно профильное, наличие базового уровня владения специальными методиками и технологиями', 
                                                               'Высшее профильное образование желательно, свободное владение специальными методиками и технологиями',
                                                               'Высшее профильное образование, требуются углубленные спец. знания и базовые в смежных областях',
                                                               'Высшее профильное образование, специальные знания в области разработок, необходимость ученой степени',
                                                               'Высшее профильное образование и дополнительное в области управления организацией и людьми'
                                                              ], 
                                             'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )


#Создадим критерии для подфактора "Навыки общения"
Communication_skills = pd.DataFrame({'Сriteria_name': ['Неслужебная бытовая коммуникация внутри подразделения/предприятия, практически полное отсутствие внешних контактов', 
                                                       'Обмен информацией с коллегами внутри подразделения, включая получение и предоставление информации и/или контакты с сотрудниками других подразделений; нечастые контакты за пределами компании, требующие лишь обычной вежливости', 
                                                       'Текущее взаимодействие (согласованность действий) с представителями других структурных подразделений, консультирование других сотрудников по вопросам, входящим в компетенцию сотрудника, и/или текущее взаимодействие (партнерство) в рамках установленных правил, процедур, договоров с внешними контрагентами, не требующее значительных навыков убеждения и влияния',
                                                       'Убеждение и влияние, в т.ч. в отношении подчиненных. Мотивация подчиненных для достижения поставленных целей, управление поведением других сотрудников в рабочих ситуациях, в т.ч. непосредственно не подчиненных. И/или рабочие контакты, требующие навыков влияния и убеждения во взаимоотношениях с внешними контрагентами',
                                                       'Рабочие контакты связаны с убеждением, преодолением сопротивления внешних контрагентов для достижения поставленных целей, умением вести переговоры в сложных ситуациях на любом уровне и/или заключаются в разрешении неоднозначных, спорных или конфликтных ситуаций внутри предприятия/холдинга',
                                                       'Публичное представление компании и ее интересов в государственных структурах, СМИ, решение возникающих при этом проблем любой сложности, лоббирование. Создание системы внутренней коммуникации в рамках компании, разработка и внедрение ее принципов'
                                                              ], 
                                              'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )



#Создадим критерии для подфактора "Знание бизнеса"
Business_knowledge = pd.DataFrame({'Сriteria_name': ['Не нужны специальные знания особенностей бизнеса', 
                                                     'Специальные знания особенностей бизнеса желательны', 
                                                     'Необходимы специальные знания особенностей бизнеса' 
                                                      ], 
                                             'Сriteria_score': [10, 85, 160],
                                             'Сriteria_number': [1, 2, 3]}
                                           )



#Создадим критерии для подфактора "Опыт работы"
Required_work_experience = pd.DataFrame({'Сriteria_name': ['Опыта работы не требуется', 
                                                           'Необходим опыт работы, не обязательно в данной области', 
                                                           'Требуется специальный опыт работы в данной области от года до 2-х лет',
                                                           'Требуется большой опыт работы в данной области (от 3-х лет)',
                                                           'Требуется серьезный опыт работы не только в данной области, но и в смежных областях',
                                                           'Кроме профессионального опыта необходим значительный опыт практического управления большим количеством сотрудников'
                                                              ], 
                                               'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )



#Создадим критерии для подфактора "Сложность решаемых вопросов"
Complexity_of_challenges  = pd.DataFrame({'Сriteria_name': ['Простые рутинные обязанности, выполнение которых требует применения лишь определенных приемов. Работа стандартная и не требует особого индивидуального подхода, решение вопросов в рамках детальных правил, инструкций', 
                                                            'Обычные обязанности с применением четко предписанных стандартных методов, использованием различных процедур. Решение проблемы требует выбора и применения одной из этих процедур. Выбор решения определяется предыдущим опытом и верность результата можно немедленно проверить', 
                                                            'Различные ситуации, требующие выбора решения путем применения приобретенного знания. Принятие решений базируется на прошлом опыте (в котором часто заключен правильный ответ), решение может быть быстро проверено на правильность. В работе необходимо формировать собственное мнение для принятия решений по планированию и выбору курса действий из различных имеющихся методов и приемов',
                                                            'Широкий круг обязанностей, требующий общего знания политики и методов работы компании, а также их применения в случаях, не оговоренных выше. Или обширные обязанности по сложной технической работе. Требуют самостоятельной работы по достижению общих целей, разработки новых методов или решения сложных технических проблем, трансформации или адаптации стандартных процедур к новым условиям принятия решений на основе имевших место прецедентов и политики компании. В работе необходимо проявлять инициативу, давать рекомендации по изменению методов работы',
                                                            'Сложная работа в области высоких технологий и комплексных проектов, в ходе которой возникают все новые или постоянно меняющиеся задачи. На данном уровне нет правильных ответов. Принятие решения требует отказа от существующих решений и прошлого опыта. Необходимо находить новые факты, которые позволят осуществить дальнейшее определение проблемы и найти верное решение. Принятое решение может быть управленческим, коммерческим или техническим',
                                                            'Участие в формировании и реализации политики компании, целей и программ ее главных отделов или структур. Обязанности крайне сложные и постоянно меняющиеся, часто требуют принятия решений при ограниченности или сомнительности информации. Решения принимаются при отсутствии устоявшихся правил или прецедентов, исходя из общих задач компании. Решение вопросов в рамках общей философии бизнеса, и/или научных принципов, связанных с коммерческими и гуманитарными ценностями'
                                                              ], 
                                               'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )


#Создадим критерии для подфактора "Область решаемых вопросов"
Area_of_challenges  = pd.DataFrame({'Сriteria_name': ['Совокупность операций', 
                                                      'Процесс', 
                                                      'Смежные процессы',
                                                      'Функция',
                                                      'Кросс-функциональный',
                                                      'Корпоративный'
                                                     ], 
                                               'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )


#Создадим критерии для подфактора "Самостоятельность в принятии решений"
Independence_in_decision_making = pd.DataFrame({'Сriteria_name': ['Нет необходимости принятия самостоятельных решений, действует под непосредственным руководством, с частым и регулярным контролем исполнения', 
                                                                  'Принимает стандартные решения из заданного набора в рамках прописанных функциональных инструкций при жестком контроле со стороны руководства', 
                                                                  'Принимает оптимальные решения в рамках прописанных функциональных процедур с учетом оперативной ситуации. Действует под общим руководством, обращаясь к руководителю в случае возникновения вопросов',
                                                                  'Под руководством, когда перед сотрудником ставят определенную цель, и он самостоятельно планирует и организовывает свою работу, обращаясь к руководителю только при возникновении нестандартных ситуаций',
                                                                  'Принимает решения, направленные на реализацию бизнес-целей подразделения в рамках стратегии. К руководителю обращается редко, только при необходимости разъяснения или трактовки политики компании',
                                                                  'Принимает решения, направленные на реализацию стратегических целей компании. Действия контролируются по выполнению основных финансовых показателей и достижению стратегических целей'
                                                                ], 
                                               'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )



#Создадим критерии для подфактора "Ответственность за финансовые результаты"
Responsibility_for_financial_results = pd.DataFrame({'Сriteria_name': ['Ответственность только за свою работу, нет ответственности за финансовый результат своей деятельности', 
                                                                       'Ответственность за финансовые результаты отдельных действий под контролем непосредственного руководителя', 
                                                                       'Ответственность за финансовые результаты регулярных действий в рамках функциональных обязанностей',
                                                                       'Выработка решений, приводящих к финансовым результатам рабочей группы или подразделения, согласование решений с непосредственным руководителем',
                                                                       'Полная ответственность за финансовые результаты работы подразделения, за материальные ценности, организационные расходы в рамках бюджета подразделения',
                                                                       'Полная ответственность за финансовые и иные результаты целого направления работ (группы подразделений)'
                                                                       ], 
                                               'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )


#Создадим критерии для подфактора "Ответственность за работу других"
Responsibility_for_work_of_others = pd.DataFrame({'Сriteria_name': ['Сотрудник ответственен только за свою работу.', 
                                                                    'Некоторое руководство и контроль за некоторыми операциями других сотрудников', 
                                                                    'Управление многими сотрудниками или подразделением, координация работы с другими руководителями',
                                                                    'Сотрудник ответственен за координацию группы отделов',
                                                                    'Сотрудник ответственен за координацию нескольких функций',
                                                                    'Руководство самостоятельной бизнес-единицей'
                                                                    ], 
                                              'Сriteria_score': [10, 40, 70, 100, 130, 160],
                                             'Сriteria_number': [1, 2, 3, 4, 5, 6]}
                                           )

#Создадим функцию для ввода данных
def input_partial_row():
    #Ввод данных для нужных столбцов
    Position_code = input("Введите индентификатор должности : ")
    Position = input("Введите название должности : ")
    Division = input("Введите название подразделения: ")
    Function = input("Введите название функции: ")
#выведем критерии для подфактора "Специальные знания и умения" пользователю через объект Console
#Создаем название таблицы и включаем линии между строками
    Console_Special_knowledge_and_skills = Table(title="Специальные знания и умения",show_lines=True) 
#Создаем столбец, устанавливаем цвет шрифта для строк к нем и ширину стролбца
    Console_Special_knowledge_and_skills.add_column("Номер критерия", style="cyan") 
#Создаем столбец, устанавливаем цвет шрифта для строк к нем и ширину стролбца
    Console_Special_knowledge_and_skills.add_column("Критерий",  style="grey66", width=50) 
#Вводим значения строк для столбцов
    Console_Special_knowledge_and_skills.add_row("1","Достаточно среднего или н/высшего образования,специальных знаний не требуется")
    Console_Special_knowledge_and_skills.add_row("2","Необходимо высшее образование, не обязательно профильное, наличие базового уровня владения специальными методиками и технологиями")
    Console_Special_knowledge_and_skills.add_row("3","Высшее профильное образование желательно, свободное владение специальными методиками и технологиями")
    Console_Special_knowledge_and_skills.add_row("4","Высшее профильное образование, требуются углубленные спец. знания и базовые в смежных областях")
    Console_Special_knowledge_and_skills.add_row("5","Высшее профильное образование, специальные знания в области разработок, необходимость ученой степени")
    Console_Special_knowledge_and_skills.add_row("6","Высшее профильное образование и дополнительное в области управления организацией и людьм")
#создаем объект Console
    console = Console() 
#выводим таблицу
    console.print(Console_Special_knowledge_and_skills)
#Просим пользователя ввести номер критерия из перечня в объекте Console
    Сhoice_Special_knowledge_and_skills = "неверно введены данные"
    while Сhoice_Special_knowledge_and_skills == "неверно введены данные":
        Сhoice_Special_knowledge_and_skills = int(input("введите номер критерия, который соответствует должности : "))
        action = Special_knowledge_and_skills.loc[Special_knowledge_and_skills['Сriteria_number'] == Сhoice_Special_knowledge_and_skills]
        if action.shape[0]==1:
            Сhoice_Special_knowledge_and_skills = Сhoice_Special_knowledge_and_skills
            print(Сhoice_Special_knowledge_and_skills)
            break
        else:
            Сhoice_Special_knowledge_and_skills = "неверно введены данные"
            print(Сhoice_Special_knowledge_and_skills)     
    #Навыки общения
    Console_Communication_skills = Table(title="Навыки общения",show_lines=True) 
    Console_Communication_skills.add_column("Номер критерия", style="cyan") 
    Console_Communication_skills.add_column("Критерий",  style="grey66", width=50)  
    Console_Communication_skills.add_row("1","Неслужебная бытовая коммуникация внутри подразделения/предприятия, практически полное отсутствие внешних контактов")
    Console_Communication_skills.add_row("2","Обмен информацией с коллегами внутри подразделения, включая получение и предоставление информации и/или контакты с сотрудниками других подразделений; нечастые контакты за пределами компании, требующие лишь обычной вежливости")
    Console_Communication_skills.add_row("3","Текущее взаимодействие (согласованность действий) с представителями других структурных подразделений, консультирование других сотрудников по вопросам, входящим в компетенцию сотрудника, и/или текущее взаимодействие (партнерство) в рамках установленных правил, процедур, договоров с внешними контрагентами, не требующее значительных навыков убеждения и влияния")
    Console_Communication_skills.add_row("4","Убеждение и влияние, в т.ч. в отношении подчиненных. Мотивация подчиненных для достижения поставленных целей, управление поведением других сотрудников в рабочих ситуациях, в т.ч. непосредственно не подчиненных. И/или рабочие контакты, требующие навыков влияния и убеждения во взаимоотношениях с внешними контрагентами")
    Console_Communication_skills.add_row("5","Рабочие контакты связаны с убеждением, преодолением сопротивления внешних контрагентов для достижения поставленных целей, умением вести переговоры в сложных ситуациях на любом уровне и/или заключаются в разрешении неоднозначных, спорных или конфликтных ситуаций внутри предприятия/холдинга")
    Console_Communication_skills.add_row("6","Публичное представление компании и ее интересов в государственных структурах, СМИ, решение возникающих при этом проблем любой сложности, лоббирование. Создание системы внутренней коммуникации в рамках компании, разработка и внедрение ее принципов")
    console = Console() 
    console.print(Console_Communication_skills)
    Сhoice_Special_knowledge_and_skills = "неверно введены данные"
    while Сhoice_Special_knowledge_and_skills == "неверно введены данные":
        Сhoice_Special_knowledge_and_skills = int(input("введите номер критерия, который соответствует должности : "))
        action = Communication_skills.loc[Special_knowledge_and_skills['Сriteria_number'] == Сhoice_Special_knowledge_and_skills]
        if action.shape[0]==1:
            Сhoice_Special_knowledge_and_skills = Сhoice_Special_knowledge_and_skills
            print(Сhoice_Special_knowledge_and_skills)
            break
        else:
            Сhoice_Special_knowledge_and_skills = "неверно введены данные"
            print(Сhoice_Special_knowledge_and_skills)
    #Знание бизнеса
    Console_Business_knowledge = Table(title="Знание бизнеса",show_lines=True) 
    Console_Business_knowledge.add_column("Номер критерия", style="cyan") 
    Console_Business_knowledge.add_column("Критерий",  style="grey66", width=50) 
    Console_Business_knowledge.add_row("1","Не нужны специальные знания особенностей бизнеса")
    Console_Business_knowledge.add_row("2","Специальные знания особенностей бизнеса желательны")
    Console_Business_knowledge.add_row("3","Необходимы специальные знания особенностей бизнеса")
    console = Console() 
    console.print(Console_Business_knowledge)
    Сhoice_Business_knowledge = "неверно введены данные"
    while Сhoice_Business_knowledge == "неверно введены данные":
        Сhoice_Business_knowledge = int(input("введите номер критерия, который соответствует должности : "))
        action = Business_knowledge.loc[Business_knowledge['Сriteria_number'] == Сhoice_Business_knowledge]
        if action.shape[0]==1:
            Сhoice_Business_knowledge = Сhoice_Business_knowledge
            print(Сhoice_Business_knowledge)
            break
        else:
            Сhoice_Business_knowledge = "неверно введены данные"
            print(Сhoice_Business_knowledge)
#Опыт работы
    Console_Required_work_experience = Table(title="Опыт работы",show_lines=True) 
    Console_Required_work_experience.add_column("Номер критерия", style="cyan") 
    Console_Required_work_experience.add_column("Критерий",  style="grey66", width=50)  
    Console_Required_work_experience.add_row("1","Опыта работы не требуется")
    Console_Required_work_experience.add_row("2","Необходим опыт работы, не обязательно в данной области")
    Console_Required_work_experience.add_row("3","Требуется специальный опыт работы в данной области от года до 2-х лет")
    Console_Required_work_experience.add_row("4","Требуется большой опыт работы в данной области (от 3-х лет)")
    Console_Required_work_experience.add_row("5","Требуется серьезный опыт работы не только в данной области, но и в смежных областях")
    Console_Required_work_experience.add_row("6","Кроме профессионального опыта необходим значительный опыт практического управления большим количеством сотрудников")
    console = Console() 
    console.print(Console_Required_work_experience)
    Сhoice_Required_work_experience = "неверно введены данные"
    while Сhoice_Required_work_experience == "неверно введены данные":
        Сhoice_Required_work_experience = int(input("введите номер критерия, который соответствует должности : "))
        action = Required_work_experience.loc[Required_work_experience['Сriteria_number'] == Сhoice_Required_work_experience]
        if action.shape[0]==1:
            Сhoice_Required_work_experience = Сhoice_Required_work_experience
            print(Сhoice_Required_work_experience)
            break
        else:
            Сhoice_Required_work_experience = "неверно введены данные"
            print(Сhoice_Required_work_experience)
#Сложность решаемых вопросов
    Console_Complexity_of_challenges = Table(title="Сложность решаемых вопросов",show_lines=True) 
    Console_Complexity_of_challenges.add_column("Номер критерия", style="cyan") 
    Console_Complexity_of_challenges.add_column("Критерий",  style="grey66", width=50) 
    Console_Complexity_of_challenges.add_row("1","Простые рутинные обязанности, выполнение которых требует применения лишь определенных приемов. Работа стандартная и не требует особого индивидуального подхода, решение вопросов в рамках детальных правил, инструкций")
    Console_Complexity_of_challenges.add_row("2","Обычные обязанности с применением четко предписанных стандартных методов, использованием различных процедур. Решение проблемы требует выбора и применения одной из этих процедур. Выбор решения определяется предыдущим опытом и верность результата можно немедленно проверить")
    Console_Complexity_of_challenges.add_row("3","Различные ситуации, требующие выбора решения путем применения приобретенного знания. Принятие решений базируется на прошлом опыте (в котором часто заключен правильный ответ), решение может быть быстро проверено на правильность. В работе необходимо формировать собственное мнение для принятия решений по планированию и выбору курса действий из различных имеющихся методов и приемов")
    Console_Complexity_of_challenges.add_row("4","Широкий круг обязанностей, требующий общего знания политики и методов работы компании, а также их применения в случаях, не оговоренных выше. Или обширные обязанности по сложной технической работе. Требуют самостоятельной работы по достижению общих целей, разработки новых методов или решения сложных технических проблем, трансформации или адаптации стандартных процедур к новым условиям принятия решений на основе имевших место прецедентов и политики компании. В работе необходимо проявлять инициативу, давать рекомендации по изменению методов работы")
    Console_Complexity_of_challenges.add_row("5","Сложная работа в области высоких технологий и комплексных проектов, в ходе которой возникают все новые или постоянно меняющиеся задачи. На данном уровне нет правильных ответов. Принятие решения ребует отказа от существующих решений и прошлого опыта. Необходимо находить новые факты, которые позволят осуществить дальнейшее определение проблемы и найти верное решение. Принятое решение может быть управленческим, коммерческим или техническим")
    Console_Complexity_of_challenges.add_row("6","Участие в формировании и реализации политики компании, целей и программ ее главных отделов или структур. Обязанности крайне сложные и постоянно меняющиеся, часто требуют принятия решений при ограниченности или сомнительности информации. Решения принимаются при отсутствии устоявшихся правил или прецедентов, исходя из общих задач компании. Решение вопросов в рамках общей философии бизнеса, и/или научных принципов, связанных с коммерческими и гуманитарными ценностями")
    console = Console() 
    console.print(Console_Complexity_of_challenges)
    Сhoice_Complexity_of_challenges = "неверно введены данные"
    while Сhoice_Complexity_of_challenges == "неверно введены данные":
        Сhoice_Complexity_of_challenges = int(input("введите номер критерия, который соответствует должности : "))
        action = Complexity_of_challenges.loc[Complexity_of_challenges['Сriteria_number'] == Сhoice_Complexity_of_challenges]
        if action.shape[0]==1:
            Сhoice_Complexity_of_challenges = Сhoice_Complexity_of_challenges
            print(Сhoice_Complexity_of_challenges)
            break
        else:
            Сhoice_Complexity_of_challenges = "неверно введены данные"
            print(Сhoice_Complexity_of_challenges)     
#Область решаемых вопросов
    Console_Area_of_challenges = Table(title="Область решаемых вопросов",show_lines=True) 
    Console_Area_of_challenges.add_column("Номер критерия", style="cyan") 
    Console_Area_of_challenges.add_column("Критерий",  style="grey66", width=50) 
    Console_Area_of_challenges.add_row("1","Совокупность операций")
    Console_Area_of_challenges.add_row("2","Процесс")
    Console_Area_of_challenges.add_row("3","Смежные процессы")
    Console_Area_of_challenges.add_row("4","Функция")
    Console_Area_of_challenges.add_row("5","Кросс-функциональный")
    Console_Area_of_challenges.add_row("6","Корпоративный")
    console = Console() 
    console.print(Console_Area_of_challenges)
    Сhoice_Area_of_challenges = "неверно введены данные"
    while Сhoice_Area_of_challenges == "неверно введены данные":
        Сhoice_Area_of_challenges = int(input("введите номер критерия, который соответствует должности : "))
        action = Communication_skills.loc[Special_knowledge_and_skills['Сriteria_number'] == Сhoice_Area_of_challenges]
        if action.shape[0]==1:
            Сhoice_Area_of_challenges = Сhoice_Area_of_challenges
            print(Сhoice_Area_of_challenges)
            break
        else:
            Сhoice_Area_of_challenges = "неверно введены данные"
            print(Сhoice_Area_of_challenges)  
#Самостоятельность в принятии решений
    Console_Independence_in_decision_making = Table(title="Самостоятельность в принятии решений",show_lines=True) 
    Console_Independence_in_decision_making.add_column("Номер критерия", style="cyan") 
    Console_Independence_in_decision_making.add_column("Критерий",  style="grey66", width=50) 
    Console_Independence_in_decision_making.add_row("1","Нет необходимости принятия самостоятельных решений, действует под непосредственным руководством, с частым и регулярным контролем исполнения")
    Console_Independence_in_decision_making.add_row("2","Принимает стандартные решения из заданного набора в рамках прописанных функциональных инструкций при жестком контроле со стороны руководства")
    Console_Independence_in_decision_making.add_row("3","Принимает оптимальные решения в рамках прописанных функциональных процедур с учетом оперативной ситуации. Действует под общим руководством, обращаясь к руководителю в случае возникновения вопросов")
    Console_Independence_in_decision_making.add_row("4","Под руководством, когда перед сотрудником ставят определенную цель, и он самостоятельно планирует и организовывает свою работу, обращаясь к руководителю только при возникновении нестандартных ситуаций")
    Console_Independence_in_decision_making.add_row("5","Принимает решения, направленные на реализацию бизнес-целей подразделения в рамках стратегии. К руководителю обращается редко, только при необходимости разъяснения или трактовки политики компании")
    Console_Independence_in_decision_making.add_row("6","Принимает решения, направленные на реализацию стратегических целей компании. Действия контролируются по выполнению основных финансовых показателей и достижению стратегических целей")
    console = Console() 
    console.print(Console_Independence_in_decision_making)
    Сhoice_Independence_in_decision_making = "неверно введены данные"
    while Сhoice_Independence_in_decision_making == "неверно введены данные":
        Сhoice_Independence_in_decision_making = int(input("введите номер критерия, который соответствует должности : "))
        action = Independence_in_decision_making.loc[Independence_in_decision_making['Сriteria_number'] == Сhoice_Independence_in_decision_making]
        if action.shape[0]==1:
            Сhoice_Independence_in_decision_making = Сhoice_Independence_in_decision_making
            print(Сhoice_Independence_in_decision_making)
            break
        else:
            Сhoice_Independence_in_decision_making = "неверно введены данные"
            print(Сhoice_Independence_in_decision_making)      
#Отвественность за финансовые результаты
    Console_Responsibility_for_financial_results = Table(title="Отвественность за финансовые результаты",show_lines=True) 
    Console_Responsibility_for_financial_results.add_column("Номер критерия", style="cyan") 
    Console_Responsibility_for_financial_results.add_column("Критерий",  style="grey66", width=50) 
    Console_Responsibility_for_financial_results.add_row("1","Ответственность только за свою работу, нет ответственности за финансовый результат своей деятельности")
    Console_Responsibility_for_financial_results.add_row("2","Ответственность за финансовые результаты отдельных действий под контролем непосредственного руководителя")
    Console_Responsibility_for_financial_results.add_row("3","Ответственность за финансовые результаты регулярных действий в рамках функциональных обязанностей")
    Console_Responsibility_for_financial_results.add_row("4","Выработка решений, приводящих к финансовым результатам рабочей группы или подразделения, согласование решений с непосредственным руководителем")
    Console_Responsibility_for_financial_results.add_row("5","Полная ответственность за финансовые результаты работы подразделения, за материальные ценности, организационные расходы в рамках бюджета подразделения")
    Console_Responsibility_for_financial_results.add_row("6","Полная ответственность за финансовые и иные результаты целого направления работ\группы подразделений")
    console = Console() 
    console.print(Console_Responsibility_for_financial_results)
    Сhoice_Responsibility_for_financial_results = "неверно введены данные"
    while Сhoice_Responsibility_for_financial_results == "неверно введены данные":
        Сhoice_Responsibility_for_financial_results = int(input("введите номер критерия, который соответствует должности : "))
        action = Responsibility_for_financial_results.loc[Responsibility_for_financial_results['Сriteria_number'] == Сhoice_Responsibility_for_financial_results]
        if action.shape[0]==1:
            Сhoice_Responsibility_for_financial_results = Сhoice_Responsibility_for_financial_results
            print(Сhoice_Responsibility_for_financial_results)
            break
        else:
            Сhoice_Responsibility_for_financial_results = "неверно введены данные"
            print(Сhoice_Responsibility_for_financial_results)   
#Отвественность за работу других
    Console_Responsibility_for_work_of_others = Table(title="Отвественность за работу других",show_lines=True) 
    Console_Responsibility_for_work_of_others.add_column("Номер критерия", style="cyan") 
    Console_Responsibility_for_work_of_others.add_column("Критерий",  style="grey66", width=50)  
    Console_Responsibility_for_work_of_others.add_row("1","Сотрудник ответственен только за свою работу")
    Console_Responsibility_for_work_of_others.add_row("2","Некоторое руководство и контроль за некоторыми операциями других сотрудников")
    Console_Responsibility_for_work_of_others.add_row("3","Управление многими сотрудниками или подразделением, координация работы с другими руководителями")
    Console_Responsibility_for_work_of_others.add_row("4","Сотрудник ответственен за координацию группы отделов")
    Console_Responsibility_for_work_of_others.add_row("5","Сотрудник ответственен за координацию нескольких функций")
    Console_Responsibility_for_work_of_others.add_row("6","Руководство самостоятельной бизнес-единицей (компания, предприятие, фирма)")
    console = Console() 
    console.print(Console_Responsibility_for_work_of_others)
    Сhoice_Responsibility_for_work_of_others= "неверно введены данные"
    while Сhoice_Responsibility_for_work_of_others == "неверно введены данные":
        Сhoice_Responsibility_for_work_of_others = int(input("введите номер критерия, который соответствует должности : "))
        action = Responsibility_for_work_of_others.loc[Responsibility_for_work_of_others['Сriteria_number'] == Сhoice_Responsibility_for_work_of_others]
        if action.shape[0]==1:
            Сhoice_Responsibility_for_work_of_others = Сhoice_Responsibility_for_work_of_others
            print(Сhoice_Responsibility_for_work_of_others)
            break
        else:
            Сhoice_Responsibility_for_work_of_others = "неверно введены данные"
            print(Сhoice_Responsibility_for_work_of_others)
    #записываем значения подфакторов
    Score_Special_knowledge_and_skills = Special_knowledge_and_skills.loc[Special_knowledge_and_skills['Сriteria_number'] == Сhoice_Special_knowledge_and_skills, 'Сriteria_score'].iloc[0]
    Score_Communication_skills = Communication_skills.loc[Communication_skills['Сriteria_number'] == Сhoice_Special_knowledge_and_skills, 'Сriteria_score'].iloc[0]        
    Score_Business_knowledge = Business_knowledge.loc[Business_knowledge['Сriteria_number'] == Сhoice_Business_knowledge, 'Сriteria_score'].iloc[0]
    Score_Required_work_experience = Required_work_experience.loc[Required_work_experience['Сriteria_number'] == Сhoice_Required_work_experience, 'Сriteria_score'].iloc[0]
    Score_Complexity_of_challenges = Complexity_of_challenges.loc[Complexity_of_challenges['Сriteria_number'] == Сhoice_Complexity_of_challenges, 'Сriteria_score'].iloc[0]
    Score_Area_of_challenges = Area_of_challenges.loc[Area_of_challenges['Сriteria_number'] ==Сhoice_Area_of_challenges, 'Сriteria_score'].iloc[0]
    Score_Independence_in_decision_making = Independence_in_decision_making.loc[Independence_in_decision_making['Сriteria_number'] == Сhoice_Independence_in_decision_making, 'Сriteria_score'].iloc[0]
    Score_Responsibility_for_financial_results = Responsibility_for_financial_results.loc[Responsibility_for_financial_results['Сriteria_number'] == Сhoice_Responsibility_for_financial_results, 'Сriteria_score'].iloc[0]
    Score_Responsibility_for_work_of_others = Responsibility_for_work_of_others.loc[Responsibility_for_work_of_others['Сriteria_number'] == Сhoice_Responsibility_for_work_of_others, 'Сriteria_score'].iloc[0]
    
    Name_Special_knowledge_and_skills = Special_knowledge_and_skills.loc[Special_knowledge_and_skills['Сriteria_number'] == Сhoice_Special_knowledge_and_skills, 'Сriteria_name'].iloc[0]
    Name_Communication_skills = Communication_skills.loc[Communication_skills['Сriteria_number'] == Сhoice_Special_knowledge_and_skills, 'Сriteria_name'].iloc[0]        
    Name_Business_knowledge = Business_knowledge.loc[Business_knowledge['Сriteria_number'] == Сhoice_Business_knowledge, 'Сriteria_name'].iloc[0]
    Name_Required_work_experience = Required_work_experience.loc[Required_work_experience['Сriteria_number'] == Сhoice_Required_work_experience, 'Сriteria_name'].iloc[0]
    Name_Complexity_of_challenges = Complexity_of_challenges.loc[Complexity_of_challenges['Сriteria_number'] == Сhoice_Complexity_of_challenges, 'Сriteria_name'].iloc[0]
    Name_Area_of_challenges = Area_of_challenges.loc[Area_of_challenges['Сriteria_number'] ==Сhoice_Area_of_challenges, 'Сriteria_name'].iloc[0]
    Name_Independence_in_decision_making = Independence_in_decision_making.loc[Independence_in_decision_making['Сriteria_number'] == Сhoice_Independence_in_decision_making, 'Сriteria_name'].iloc[0]
    Name_Responsibility_for_financial_results = Responsibility_for_financial_results.loc[Responsibility_for_financial_results['Сriteria_number'] == Сhoice_Responsibility_for_financial_results, 'Сriteria_name'].iloc[0]
    Name_Responsibility_for_work_of_others = Responsibility_for_work_of_others.loc[Responsibility_for_work_of_others['Сriteria_number'] == Сhoice_Responsibility_for_work_of_others, 'Сriteria_name'].iloc[0]
    
    #рассчитываем итоговое количество баллов для должности
    Total_Score = round((50/150)*((10/50)*int(Score_Special_knowledge_and_skills)+(10/50)*int(Score_Communication_skills)+(10/50)*int(Score_Business_knowledge)+(20/50)*int(Score_Required_work_experience))+(50/150)*((30/60)*int(Score_Complexity_of_challenges)+(30/60)*int(Score_Area_of_challenges))+(50/150)*((30/70)*int(Score_Independence_in_decision_making)+(20/70)*int(Score_Responsibility_for_financial_results)+(20/70)*int(Score_Responsibility_for_work_of_others)))
    #опеределяем грейд
    Grade = round(int(Total_Score)/10)
    #Создаём строку с 15 столбцами, заполняя остальные пустыми строками
    row = ["" for _ in range(24)]
    row = {'Position_code': Position_code, 
           'Position': Position, 
           'Division': Division,
           'Function': Function,
           'Score_Special_knowledge_and_skills': Score_Special_knowledge_and_skills,
           'Name_Special_knowledge_and_skills': Name_Special_knowledge_and_skills,
           'Score_Communication_skills': Score_Communication_skills,
           'Name_Communication_skills': Name_Communication_skills,
           'Score_Business_knowledge': Score_Business_knowledge,
           'Name_Business_knowledge': Name_Business_knowledge,
           'Score_Required_work_experience': Score_Required_work_experience,
           'Name_Required_work_experience': Name_Required_work_experience,
           'Score_Complexity_of_challenges': Score_Complexity_of_challenges,
           'Name_Complexity_of_challenges': Name_Complexity_of_challenges,
           'Score_Area_of_challenges': Score_Area_of_challenges,
           'Name_Area_of_challenges': Name_Area_of_challenges,
           'Score_Independence_in_decision_making': Score_Independence_in_decision_making,
           'Name_Independence_in_decision_making': Name_Independence_in_decision_making,
           'Score_Responsibility_for_financial_results': Score_Responsibility_for_financial_results,
           'Name_Responsibility_for_financial_results': Name_Responsibility_for_financial_results,
           'Score_Responsibility_for_work_of_others': Score_Responsibility_for_work_of_others,
           'Name_Responsibility_for_work_of_others': Name_Responsibility_for_work_of_others,
           'Total_Score': Total_Score,
           'Grade': Grade}
    return row                

#Создадим функцию main
def main():
    table = []
    n = int(input("Сколько строк хотите ввести? "))
    for i in range(n):
        row = input_partial_row()
        table.append(row)
        List_of_evaluated_positions = pd.DataFrame(table)
        # Функция для объединения строк через ; и с нового абзаца
        def join_texts(values):
            return ';\n\n'.join(str(v) for v in values if pd.notnull(v))
        # Создание матрицы с помощью pivot_table
        matrix = pd.pivot_table(
        List_of_evaluated_positions,
        index='Grade',
        columns='Function',
        values='Position',
        aggfunc=join_texts,
        fill_value=''
        )
 
    # Выводим исходный DataFrame
        print("Исходный DataFrame:")
        print(List_of_evaluated_positions)
 
    # Выводим результирующую матрицу
        print("\nРезультирующая матрица:")
        print(matrix)
        with pd.ExcelWriter('D:\Грейдинг\output.xlsx') as writer:
            List_of_evaluated_positions.to_excel(writer, sheet_name='grade_calculatior', index=False)
            matrix.to_excel(writer, sheet_name='matrix')
    print("\nВведённая таблица:")
    print(List_of_evaluated_positions)
if __name__ == "__main__":
    main()
