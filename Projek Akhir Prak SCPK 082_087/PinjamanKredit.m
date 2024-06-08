function varargout = PinjamanKredit(varargin)
% PINJAMANKREDIT MATLAB code for PinjamanKredit.fig
%      PINJAMANKREDIT, by itself, creates a new PINJAMANKREDIT or raises the existing
%      singleton*.
%
%      H = PINJAMANKREDIT returns the handle to a new PINJAMANKREDIT or the handle to
%      the existing singleton*.
%
%      PINJAMANKREDIT('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PINJAMANKREDIT.M with the given input arguments.
%
%      PINJAMANKREDIT('Property','Value',...) creates a new PINJAMANKREDIT or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before PinjamanKredit_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to PinjamanKredit_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help PinjamanKredit

% Last Modified by GUIDE v2.5 27-May-2024 10:13:07

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @PinjamanKredit_OpeningFcn, ...
                   'gui_OutputFcn',  @PinjamanKredit_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before PinjamanKredit is made visible.
function PinjamanKredit_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to PinjamanKredit (see VARARGIN)

% Choose default command line output for PinjamanKredit
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes PinjamanKredit wait for user response (see UIRESUME)
% uiwait(handles.figure1);




% --- Outputs from this function are returned to the command line.
function varargout = PinjamanKredit_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in showdataButton.
function showdataButton_Callback(hObject, eventdata, handles)
% hObject    handle to showdataButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% FUNCTION UNTUK MENAMPILKAN DATA DARI EXCEL KE TABEL GUI MATLAB

try
    % Deteksi opsi impor dari file Excel 'CreditScoreData.xlsx'
    opts = detectImportOptions('CreditScoreData.xlsx');

    % Pilih kolom yang diinginkan (kolom 4, 9, 10, 11, 12, 17, 19, 20, 22, 24)
    opts.SelectedVariableNames = opts.VariableNames([4, 9, 10, 11, 12, 17, 19, 20, 22, 24]);

    % Baca tabel dari file Excel ke dalam tabel dengan opsi yang telah ditentukan
    dataTable = readtable('CreditScoreData.xlsx', opts);

    % Konversi tabel ke format sel agar sesuai dengan format uitable
    dataCell = table2cell(dataTable);

    % Set data ke dalam uitable pada GUI
    set(handles.uitable1, 'data', dataCell);
catch
    msgbox('File Tidak Ada', 'Error', 'error'); % Menampilkan pesan error jika file tidak ada
end



function monthly_inhand_salary_Callback(hObject, eventdata, handles)
% hObject    handle to monthly_inhand_salary (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of monthly_inhand_salary as text
%        str2double(get(hObject,'String')) returns contents of monthly_inhand_salary as a double


% --- Executes during object creation, after setting all properties.
function monthly_inhand_salary_CreateFcn(hObject, eventdata, handles)
% hObject    handle to monthly_inhand_salary (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function num_bank_account_Callback(hObject, eventdata, handles)
% hObject    handle to num_bank_account (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of num_bank_account as text
%        str2double(get(hObject,'String')) returns contents of num_bank_account as a double


% --- Executes during object creation, after setting all properties.
function num_bank_account_CreateFcn(hObject, eventdata, handles)
% hObject    handle to num_bank_account (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function num_credit_card_Callback(hObject, eventdata, handles)
% hObject    handle to num_credit_card (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of num_credit_card as text
%        str2double(get(hObject,'String')) returns contents of num_credit_card as a double


% --- Executes during object creation, after setting all properties.
function num_credit_card_CreateFcn(hObject, eventdata, handles)
% hObject    handle to num_credit_card (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function interest_rate_Callback(hObject, eventdata, handles)
% hObject    handle to interest_rate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of interest_rate as text
%        str2double(get(hObject,'String')) returns contents of interest_rate as a double


% --- Executes during object creation, after setting all properties.
function interest_rate_CreateFcn(hObject, eventdata, handles)
% hObject    handle to interest_rate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function change_credit_limit_Callback(hObject, eventdata, handles)
% hObject    handle to change_credit_limit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of change_credit_limit as text
%        str2double(get(hObject,'String')) returns contents of change_credit_limit as a double


% --- Executes during object creation, after setting all properties.
function change_credit_limit_CreateFcn(hObject, eventdata, handles)
% hObject    handle to change_credit_limit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function credit_mix_Callback(hObject, eventdata, handles)
% hObject    handle to credit_mix (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of credit_mix as text
%        str2double(get(hObject,'String')) returns contents of credit_mix as a double


% --- Executes during object creation, after setting all properties.
function credit_mix_CreateFcn(hObject, eventdata, handles)
% hObject    handle to credit_mix (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function outstanding_debt_Callback(hObject, eventdata, handles)
% hObject    handle to outstanding_debt (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of outstanding_debt as text
%        str2double(get(hObject,'String')) returns contents of outstanding_debt as a double


% --- Executes during object creation, after setting all properties.
function outstanding_debt_CreateFcn(hObject, eventdata, handles)
% hObject    handle to outstanding_debt (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function credit_history_age_months_Callback(hObject, eventdata, handles)
% hObject    handle to credit_history_age_months (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of credit_history_age_months as text
%        str2double(get(hObject,'String')) returns contents of credit_history_age_months as a double


% --- Executes during object creation, after setting all properties.
function credit_history_age_months_CreateFcn(hObject, eventdata, handles)
% hObject    handle to credit_history_age_months (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function total_emi_per_month_Callback(hObject, eventdata, handles)
% hObject    handle to total_emi_per_month (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of total_emi_per_month as text
%        str2double(get(hObject,'String')) returns contents of total_emi_per_month as a double


% --- Executes during object creation, after setting all properties.
function total_emi_per_month_CreateFcn(hObject, eventdata, handles)
% hObject    handle to total_emi_per_month (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in resultButton.
function resultButton_Callback(hObject, eventdata, handles)
% hObject    handle to resultButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% FUNCTION UNTUK MENGOLAH DATA INPUT USER AGAR DIDAPATKAN 10 NASABAH
% TERBAIK DAN 1 TERBAIK

% Error handling untuk file Excel jika filenya tidak ada
try
    data = readtable('CreditScoreData.xlsx'); % Membaca data dari file Excel 'CreditScoreData.xlsx' ke dalam tabel
catch
    msgbox('File Tidak Ada', 'Error','error'); % Menampilkan pesan error jika file tidak ada
    return;
end

% User Input (konversi ke tipe data numerik)
fields = {'monthly_inhand_salary', 'num_bank_account', 'num_credit_card', 'interest_rate', 'change_credit_limit', 'credit_mix', 'outstanding_debt', 'credit_history_age_months', 'total_emi_per_month'};
w = zeros(1, numel(fields)); % Inisialisasi vektor bobot dengan ukuran yang sesuai dengan jumlah field

for i = 1:numel(fields)
    userInput = get(handles.(fields{i}), 'String'); % Mengambil input pengguna dari GUI
    if isempty(userInput)
        msgbox('Inputan masih ada yang kosong', 'Error', 'error'); % Menampilkan pesan error jika ada input yang kosong
        return;
    end
    
    w(i) = str2double(userInput); % Konversi input dari string ke double
    if isnan(w(i)) || w(i) < 1 || w(i) > 9
        msgbox('Inputan Tidak Dalam Range 1-9', 'Error', 'error'); % Menampilkan pesan error jika input tidak valid
        return;
    end
end

% Inisialisasi rating kecocokan, atribut, dan nilai bobot antar kriteria

try
    % Menyusun matriks data rating kecocokan masing-masing alternatif (diambil
    % dari kolom kriteria yang sesuai di excel)
    x = [data.Monthly_Inhand_Salary, data.Num_Bank_Accounts, data.Num_Credit_Card, ...
        data.Interest_Rate, data.Changed_Credit_Limit, data.Credit_Mix, ...
        data.Outstanding_Debt, data.Credit_History_Age_Months, data.Total_EMI_per_month];

    % Menentukan jenis atribut tiap-tiap kriteria yaitu 1 = benefit dan 0 = cost
    k = [1 1 0 0 1 1 0 1 0];

    % Menyusun vektor bobot (weight) berdasarkan input pengguna
    w = w ./ sum(w); % Normalisasi bobot dengan membagi bobot per kriteria dengan jumlah total seluruh bobot (jumlah total bobot menjadi 1)

    % Tahapan pertama, perbaikan bobot
    [m, n] = size(x);  % Inisialisasi ukuran x
    for j = 1:n
        if k(j) == 0
            w(j) = -1 * w(j); % Mengubah bobot menjadi negatif untuk atribut cost
        end
    end

    % Tahapan kedua, melakukan perhitungan vektor(S) per baris (alternatif)
    S = zeros(m, 1);
    for i = 1:m
        S(i) = prod(x(i, :) .^ w); % Menghitung nilai vektor S untuk setiap baris (alternatif) menggunakan bobot
    end

    % Tahapan ketiga, proses perangkingan
    V = S / sum(S); % Menghitung nilai vektor V untuk setiap alternatif

    % Mengurutkan Data
    [sortedV, idx] = sort(V, 'descend'); % Mengurutkan nilai V secara menurun dan mendapatkan indeksnya
    top10 = data(idx(1:10), :); % Memilih 10 alternatif terbaik berdasarkan nilai V
    top10_selected = top10(:, [2, 4]); % Memilih kolom yang akan ditampilkan dalam uitable2 (kolom 2 dan 4)

    % Mendapatkan Nama dan Nilai Preferensi (Nilai Vektor V) nasabah terbaik
    namaNasabahTerbaik = top10.Name{1}; % Mendapatkan nama nasabah terbaik (nilai V tertinggi)
    nilaiVNasabahTerbaik = sortedV(1); % Mendapatkan nilai V tertinggi
    formattedVNasabahTerbaik = sprintf('%.8f', nilaiVNasabahTerbaik); % Memformat nilai V dengan 8 digit desimal

    % Menampilkan hasil
    set(handles.hasil1, 'String', namaNasabahTerbaik); % Menampilkan nama nasabah terbaik di hasil1
    set(handles.hasil2, 'String', formattedVNasabahTerbaik); % Menampilkan nilai V nasabah terbaik di hasil2
    set(handles.uitable2, 'Data', table2cell(top10_selected)); % Menampilkan data 10 nasabah terbaik di uitable2

    % Menampilkan pesan bahwa perhitungan WP selesai
    msgbox('Perhitungan WP selesai', 'Success', 'help'); % Menampilkan pesan success jika perhitungan berhasil dan selesai

catch
    msgbox('Terjadi kesalahan dalam perhitungan', 'Error', 'error'); % Menampilkan pesan error jika terjadi kesalahan dalam perhitungan
end

% --- If Enable == 'on', executes on mouse press in 5 pixel border.
% --- Otherwise, executes on mouse press in 5 pixel border or over showdataButton.
function showdataButton_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to showdataButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- If Enable == 'on', executes on mouse press in 5 pixel border.
% --- Otherwise, executes on mouse press in 5 pixel border or over hasil1.
function hasil1_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to hasil1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in resetButton.
function resetButton_Callback(hObject, eventdata, handles)
% hObject    handle to resetButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% FUNCTION UNTUK MENGHAPUS / MERESET DATA INPUTAN DAN LAIN-LAIN

% Menghapus data dari uitable1
data = cell(1,10);  % Membuat sel kosong dengan ukuran 1x10
set(handles.uitable1, 'Data', data); % Mengatur data kosong ke uitable1

% Menghapus data dari uitable2
data = cell(1,2); % Membuat sel kosong dengan ukuran 1x2
set(handles.uitable2, 'Data', data); % Mengatur data kosong ke uitable2

% Mengatur ulang inputan pengguna ke nilai kosong

set(handles.monthly_inhand_salary, 'String', ''); % Mengosongkan input monthly_inhand_salary
set(handles.num_bank_account, 'String', ''); % Mengosongkan input num_bank_account
set(handles.num_credit_card, 'String', ''); % Mengosongkan input num_credit_card
set(handles.interest_rate, 'String', ''); % Mengosongkan input interest_rate
set(handles.change_credit_limit, 'String', ''); % Mengosongkan input change_credit_limit
set(handles.credit_mix, 'String', ''); % Mengosongkan input credit_mix
set(handles.outstanding_debt, 'String', ''); % Mengosongkan input outstanding_debt
set(handles.credit_history_age_months, 'String', ''); % Mengosongkan input credit_history_age_months
set(handles.total_emi_per_month, 'String', ''); % Mengosongkan input total_emi_per_month
set(handles.hasil1, 'String', ''); % Mengosongkan output hasil1 (nama nasabah terbaik)
set(handles.hasil2, 'String', ''); % Mengosongkan output hasil2 (nilai V nasabah terbaik)


function hasil1_Callback(hObject, eventdata, handles)
% hObject    handle to hasil1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of hasil1 as text
%        str2double(get(hObject,'String')) returns contents of hasil1 as a double


% --- Executes during object creation, after setting all properties.
function hasil1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to hasil1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function hasil2_Callback(hObject, eventdata, handles)
% hObject    handle to hasil2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of hasil2 as text
%        str2double(get(hObject,'String')) returns contents of hasil2 as a double


% --- Executes during object creation, after setting all properties.
function hasil2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to hasil2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
