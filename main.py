from tkinter.filedialog import askopenfilename
import pyodbc, configparser, os, csv, logging, time
from tkinter import *

logging.basicConfig(filename="C://temp/logs/log_import_{}.log".format(time.strftime("%Y%m%d-%H%M%S")),
                    format='%(asctime)s %(message)s',
                    level=logging.DEBUG,
                    filemode='w')

vLocal = os.path.dirname(__file__)

# Dados para conexao
vDadosConfig = configparser.ConfigParser()
vDadosConfig._interpolation = configparser.ExtendedInterpolation()
vDadosConfig.read(vLocal + '/config.ini')
vDadosConfig.sections()

vConnectionString = str(vDadosConfig.get('parametros', 'connectionString'))

vTabela = 'TableName'

# criação da janela principal
root = Tk()
root.title('IMPORTADOR DE ARQUIVOS PARA BD')
root.iconbitmap('file_icon.ico')
root.geometry('600x100')
root.resizable(False, False)
root.configure(bg='white')
initial_text = Label(root, text='Selecione o arquivo para importar (formato csv):',
                     font=('Helvetica', 12, 'bold'),
                     bg='white',
                     padx=20)
initial_text.grid(column=0, row=0, columnspan=1)

my_entry = Entry(root,
                 font=('Helvetica', 10, 'bold'),
                 bd=0,
                 width='50',
                 relief='solid')
my_entry_total = Entry(root,
                 font=('Helvetica', 10, 'bold'),
                 bd=0,
                 width='50',
                 relief='solid')


def copy_text():
    txt = my_entry.get()
    root.clipboard_append(txt)


def file():
    errors = list()
    type_error = list()
    filename = askopenfilename(initialdir='/', title='Selecione o arquivo csv',
                               filetypes=(('csv files', '*.csv'), ('all files', '*.*')))
    cont = 0
    localmachine = "'" + os.environ['COMPUTERNAME'] + "'"
    my_entry.delete(0, 'end')
    my_entry_total.delete(0, 'end')
    if (filename) and filename.endswith(('.csv')):
        file_path.set(filename)
        try:
            cnxn = pyodbc.connect(vConnectionString)
            cursor = cnxn.cursor()
        except Exception as e:
            logging.error("Exception occurred", exc_info=True)
            #raise e
        else:
            with open(filename, 'r') as vCSV:
                vCSVLinhas = csv.reader(vCSV, delimiter=';')
                next(vCSVLinhas)
                for vLinha in vCSVLinhas:
                    vInsertLinha = ''
                    vControle = 0
                    if len(vLinha) == 3:
                        for vColuna in range(0, len(vLinha)):
                            if vColuna == 0:
                                vEmpregado = int(vLinha[0])
                            elif vColuna == 1:
                                vDataColunastring = vLinha[1]
                            elif vColuna == 2:
                                vVales = int(vLinha[2])

                    else:
                        type_error.append('Layout incompatível!')
                        break

                    try:
                        vDataColuna = str(vDataColunastring).split("/")
                        vData = "'" + str("%s-%s-%s" % (vDataColuna[2], vDataColuna[1], vDataColuna[0])) + "'"
                        vInsertLinha = str(f'INSERT INTO {vTabela} VALUES( {vEmpregado}, {vData}, {vVales}, {localmachine},getdate())')
                        cursor.execute(vInsertLinha)
                        cont += 1

                    except (pyodbc.IntegrityError, pyodbc.DataError, IndexError) as e:
                        employee = str('%s-%s' % (vEmpregado,vDataColunastring))
                        errors.append(employee)
                        logging.error("Exception occurred", exc_info=True)
                        # raise e


                    cnxn.commit()
                cnxn.close()

            if not errors and type_error:
                my_entry.grid(column=0, row=2)
                my_entry.insert(1, f'{type_error}')
                my_entry_total.grid(column=0, row = 3)
                my_entry_total.insert(1, f'Total de linhas importadas {cont}')
            elif errors and not type_error:
                my_entry.grid(column=0, row=2)
                my_entry.insert(1, f'Erro nos registros: {errors}.')
                my_entry_total.grid(column=0, row = 3)
                my_entry_total.insert(1, f'Total de linhas importadas {cont}')
            elif errors and type_error:
                my_entry.grid(column=0, row=2)
                my_entry.insert(1, f'{type_error} e Chaves duplicadas: {errors}.')
                my_entry_total.grid(column=0, row = 3)
                my_entry_total.insert(1, f'Total de linhas importadas {cont}')

            else:
                my_entry.grid(column=0, row=2)
                my_entry.insert(1, f'Linhas importadas {cont}')

    else:
        my_entry.grid(column=0, row=2)
        my_entry.insert(1, f'Erro! Importe um arquivo .csv!')
    copy = Button(root,
                  text='Copiar erros',
                  command=copy_text(),
                  width='10',
                  relief='solid',
                  font=('Helvetica', 8, 'bold'))
    copy.grid(column=1, row=2)



button = Button(root, text='Open', command=file, width='10', relief='solid', font=('Helvetica', 8, 'bold'))
button.grid(column=1, row=1)
file_path = StringVar()
box = Label(root,
            textvariable=file_path,
            relief='solid',
            width='60',
            font=('Helvetica', 8),
            bg='white')
box.grid(column=0, row=1, sticky='EW')

file_path.set("Path")

root.mainloop()
