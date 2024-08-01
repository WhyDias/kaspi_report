package main

import (
	"encoding/csv"
	"fmt"
	"github.com/go-telegram-bot-api/telegram-bot-api/v5"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
)

var (
	costMap        map[string]float64
	nameMap        map[string]string
	categoryMap    map[string]string
	subcategoryMap map[string]string
	isCostFile     bool
	isReportFile   bool
)

// Функция для загрузки себестоимости товаров из CSV файла
func loadProductCosts(filename string) (map[string]float64, map[string]string, map[string]string, map[string]string, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, nil, nil, nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	reader.LazyQuotes = true
	records, err := reader.ReadAll()
	if err != nil {
		return nil, nil, nil, nil, err
	}

	costs := make(map[string]float64)
	names := make(map[string]string)
	categories := make(map[string]string)
	subcategories := make(map[string]string)
	for _, record := range records[1:] { // Пропускаем заголовок
		if len(record) < 5 {
			return nil, nil, nil, nil, fmt.Errorf("неверное количество полей в строке: %v", record)
		}
		log.Printf("Обработка записи: %v", record)
		product := record[0]
		mainName := record[1]
		category := record[2]
		subcategory := record[3]
		cost, err := parseCost(record[4])
		if err != nil {
			return nil, nil, nil, nil, err
		}
		costs[product] = cost
		names[product] = mainName
		categories[product] = category
		subcategories[product] = subcategory
	}

	return costs, names, categories, subcategories, nil
}

// Функция для удаления запятых и преобразования строки в float64
func parseFloat(s string) (float64, error) {
	s = strings.ReplaceAll(s, ",", "")
	return strconv.ParseFloat(s, 64)
}

// Функция для преобразования строки стоимости в float64
func parseCost(s string) (float64, error) {
	s = strings.ReplaceAll(s, "\u00a0", "") // Удаляем неразрывные пробелы
	s = strings.ReplaceAll(s, " ", "")      // Удаляем обычные пробелы
	s = strings.ReplaceAll(s, ".", "")      // Удаляем точки (разделители тысяч)
	s = strings.Replace(s, ",", ".", 1)     // Заменяем запятую на точку (десятичная точка)
	return strconv.ParseFloat(s, 64)
}

// Функция для создания новой таблицы
func createNewTable(filePath string, costMap map[string]float64, nameMap map[string]string, categoryMap map[string]string, subcategoryMap map[string]string) (string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return "", fmt.Errorf("не удалось открыть файл: %v", err)
	}
	defer f.Close()

	// Создаем новую Excel таблицу
	newFile := excelize.NewFile()
	sheetName := newFile.GetSheetName(0)

	// Добавляем заголовки новых столбцов
	headers := []string{"Название в отчете", "Категория", "Подкатегория", "Детали покупки", "Количество", "Сумма операции(тг)", "Комиссия за операцию(%)", "Комиссия за операцию(тг)", "Комиссия Kaspi Pay(тг)", "Комиссия за Доставку(тг)", "Себестоимость(тг)", "Маржа(тг)", "Валовая прибыль(тг)"}
	for i, header := range headers {
		cellIndex, _ := excelize.CoordinatesToCellName(i+1, 1)
		newFile.SetCellValue(sheetName, cellIndex, header)
	}

	var totalProfit float64
	re := regexp.MustCompile(`,?\s*(\d+)\s*шт\.?$`)

	// Заполняем новую таблицу данными
	rowIndex := 2
	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		return "", fmt.Errorf("не удалось прочитать строки: %v", err)
	}

	for _, row := range rows[7:] {
		if strings.TrimSpace(row[12]) != "Покупка" {
			continue
		}

		productCell := row[31]
		matches := re.FindStringSubmatch(productCell)
		quantity := 1
		if len(matches) > 1 {
			quantity, err = strconv.Atoi(matches[1])
			if err != nil {
				quantity = 1
			}
			productCell = strings.TrimSpace(strings.TrimSuffix(productCell, matches[0]))
		}

		mainName, ok := nameMap[productCell]
		if !ok {
			mainName = productCell
		}

		category := categoryMap[productCell]
		subcategory := subcategoryMap[productCell]

		totalRevenue, err1 := parseFloat(row[18])
		commissionPercent := row[21] // Комиссия за операцию (%) оставляем как строку
		commissionAmount, err2 := parseFloat(row[20])
		kaspiPayCommission, err3 := parseFloat(row[26])
		deliveryCommission := 0.0
		if row[29] != "" {
			deliveryCommission, err = parseFloat(row[29])
		}

		// Проверка на ошибки при преобразовании
		if err1 != nil || err2 != nil || err3 != nil || (row[29] != "" && err != nil) {
			log.Printf("Ошибка преобразования строки: %v, %v, %v, %v", err1, err2, err3, err)
			continue
		}

		cost := costMap[productCell] * float64(quantity)
		margin := totalRevenue - cost
		grossProfit := totalRevenue - cost + commissionAmount + kaspiPayCommission + deliveryCommission

		log.Printf("Заполнение строки %d: Название: %s, Категория: %s, Подкатегория: %s, Количество: %d, Сумма операции: %.2f", rowIndex, mainName, category, subcategory, quantity, totalRevenue)

		newFile.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), mainName)
		newFile.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), category)
		newFile.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), subcategory)
		newFile.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex), productCell)
		newFile.SetCellValue(sheetName, fmt.Sprintf("E%d", rowIndex), quantity)
		newFile.SetCellValue(sheetName, fmt.Sprintf("F%d", rowIndex), totalRevenue)
		newFile.SetCellValue(sheetName, fmt.Sprintf("G%d", rowIndex), commissionPercent) // как строка
		newFile.SetCellValue(sheetName, fmt.Sprintf("H%d", rowIndex), commissionAmount)
		newFile.SetCellValue(sheetName, fmt.Sprintf("I%d", rowIndex), kaspiPayCommission)
		newFile.SetCellValue(sheetName, fmt.Sprintf("J%d", rowIndex), deliveryCommission)
		newFile.SetCellValue(sheetName, fmt.Sprintf("K%d", rowIndex), cost)
		newFile.SetCellValue(sheetName, fmt.Sprintf("L%d", rowIndex), margin)
		newFile.SetCellValue(sheetName, fmt.Sprintf("M%d", rowIndex), grossProfit)

		totalProfit += grossProfit
		rowIndex++
	}

	// Сохраняем новую таблицу
	newFilePath := strings.Replace(filePath, ".xlsx", "_new.xlsx", 1)
	if err := newFile.SaveAs(newFilePath); err != nil {
		return "", fmt.Errorf("не удалось сохранить файл: %v", err)
	}

	log.Printf("Новая таблица сохранена по пути: %s", newFilePath)
	return newFilePath, nil
}

func main() {
	bot, err := tgbotapi.NewBotAPI("7386513840:AAFzi2ixblH3DGeWTRoP0cURP6XP0A-ml3M")
	if err != nil {
		log.Panic(err)
	}

	bot.Debug = true

	log.Printf("Авторизован на аккаунте %s", bot.Self.UserName)

	u := tgbotapi.NewUpdate(0)
	u.Timeout = 60

	updates := bot.GetUpdatesChan(u)

	for update := range updates {
		if update.Message == nil { // игнорируем любые не Message обновления
			continue
		}

		if update.Message.IsCommand() {
			switch update.Message.Command() {
			case "start":
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Привет! Отправьте мне файл себестоимости товаров (product_costs.csv).")
				bot.Send(msg)
				isCostFile = false
				isReportFile = false
			default:
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Я не знаю такой команды.")
				bot.Send(msg)
			}
		}

		if update.Message.Document != nil {
			file := update.Message.Document
			filePath := fmt.Sprintf("/tmp/%s", file.FileName)
			fileURL, err := bot.GetFileDirectURL(file.FileID)
			if err != nil {
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при получении файла: %v", err))
				bot.Send(msg)
				log.Printf("Ошибка при получении файла: %v", err)
				continue
			}

			log.Printf("Скачивание файла: %s", filePath)

			// Скачиваем файл
			err = downloadFile(filePath, fileURL)
			if err != nil {
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при скачивании файла: %v", err))
				bot.Send(msg)
				log.Printf("Ошибка при скачивании файла: %v", err)
				continue
			}

			if strings.HasSuffix(file.FileName, ".csv") && !isCostFile {
				// Загрузка себестоимости товаров
				log.Printf("Загрузка файла себестоимости товаров: %s", filePath)
				costMap, nameMap, categoryMap, subcategoryMap, err = loadProductCosts(filePath)
				if err != nil {
					msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Не удалось загрузить себестоимость товаров: %v", err))
					bot.Send(msg)
					log.Printf("Не удалось загрузить себестоимость товаров: %v", err)
					continue
				}
				isCostFile = true
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Файл себестоимости товаров загружен. Теперь отправьте Excel отчет для обработки.")
				bot.Send(msg)
			} else if strings.HasSuffix(file.FileName, ".xlsx") && isCostFile {
				// Обработка файла
				log.Printf("Обработка Excel отчета: %s", filePath)
				newFilePath, err := createNewTable(filePath, costMap, nameMap, categoryMap, subcategoryMap)
				if err != nil {
					msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при обработке файла: %v", err))
					bot.Send(msg)
					log.Printf("Ошибка при обработке файла: %v", err)
					continue
				}

				// Отправка обработанного файла обратно
				log.Printf("Отправка обработанного файла: %s", newFilePath)
				doc := tgbotapi.NewDocument(update.Message.Chat.ID, tgbotapi.FilePath(newFilePath))
				doc.Caption = "Обработанный файл"
				_, err = bot.Send(doc)
				if err != nil {
					msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при отправке файла: %v", err))
					bot.Send(msg)
					log.Printf("Ошибка при отправке файла: %v", err)
				}
				isReportFile = true
			} else {
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Пожалуйста, сначала отправьте файл себестоимости товаров (product_costs.csv), а затем Excel отчет (.xlsx).")
				bot.Send(msg)
				log.Printf("Получен неправильный файл или порядок файлов нарушен.")
			}
		}
	}
}

// Функция для скачивания файла по URL
func downloadFile(filepath string, url string) error {
	resp, err := http.Get(url)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	out, err := os.Create(filepath)
	if err != nil {
		return err
	}
	defer out.Close()

	_, err = io.Copy(out, resp.Body)
	return err
}
