package main

//
//import (
//	"encoding/csv"
//	"fmt"
//	"io"
//	"log"
//	"net/http"
//	"os"
//	"regexp"
//	"strconv"
//	"strings"
//
//	"github.com/go-telegram-bot-api/telegram-bot-api/v5"
//	"github.com/xuri/excelize/v2"
//)
//
//var (
//	costMap      map[string]float64
//	isCostFile   bool
//	isReportFile bool
//)
//
//// Функция для загрузки себестоимости товаров из CSV файла
//func loadProductCosts(filename string) (map[string]float64, error) {
//	file, err := os.Open(filename)
//	if err != nil {
//		return nil, err
//	}
//	defer file.Close()
//
//	reader := csv.NewReader(file)
//	reader.LazyQuotes = true
//	records, err := reader.ReadAll()
//	if err != nil {
//		return nil, err
//	}
//
//	costs := make(map[string]float64)
//	for _, record := range records[1:] { // Пропускаем заголовок
//		if len(record) < 2 {
//			return nil, fmt.Errorf("неверное количество полей в строке: %v", record)
//		}
//		product := record[0]
//		cost, err := strconv.ParseFloat(record[1], 64)
//		if err != nil {
//			return nil, err
//		}
//		costs[product] = cost
//	}
//
//	return costs, nil
//}
//
//// Функция для удаления запятых и преобразования строки в float64
////func parseFloat(s string) (float64, error) {
////	s = strings.ReplaceAll(s, ",", "")
////	return strconv.ParseFloat(s, 64)
////}
//
//// Функция для обработки Excel файла
//func processExcel(filePath string, costMap map[string]float64) (string, error) {
//	f, err := excelize.OpenFile(filePath)
//	if err != nil {
//		return "", fmt.Errorf("не удалось открыть файл: %v", err)
//	}
//	defer f.Close()
//
//	// Создаем новый Excel файл
//	newFile := excelize.NewFile()
//	sheetName := newFile.GetSheetName(0)
//
//	// Читаем строки из исходного файла
//	rows, err := f.GetRows(f.GetSheetName(0))
//	if err != nil {
//		return "", fmt.Errorf("не удалось прочитать строки: %v", err)
//	}
//
//	// Копируем строки до 6-й включительно
//	for i, row := range rows {
//		for j, cell := range row {
//			cellIndex, _ := excelize.CoordinatesToCellName(j+1, i+1)
//			newFile.SetCellValue(sheetName, cellIndex, cell)
//		}
//	}
//
//	// Добавляем заголовки новых столбцов в 7-ю строку
//	newFile.SetCellValue(sheetName, "AH7", "Себестоимость")
//	newFile.SetCellValue(sheetName, "AI7", "Чистая Прибыль")
//
//	var totalProfit float64
//
//	// Регулярное выражение для поиска количества штук в названии товара
//	re := regexp.MustCompile(`,?\s*(\d+)\s*шт\.?$`)
//
//	// Заполняем новые столбцы данными начиная с 8-й строки
//	for i := 7; i < len(rows); i++ {
//		// Проверка типа операции в столбце M
//		if strings.TrimSpace(rows[i][12]) != "Покупка" {
//			continue
//		}
//
//		productCell := rows[i][31] // Название товара в колонке AF (31-я колонка, так как индексация начинается с 0)
//		matches := re.FindStringSubmatch(productCell)
//		quantity := 1
//		productName := productCell
//		if len(matches) > 1 {
//			quantity, err = strconv.Atoi(matches[1])
//			if err != nil {
//				quantity = 1 // Если не удается преобразовать, по умолчанию 1
//			}
//			productName = strings.TrimSpace(strings.TrimSuffix(productCell, matches[0]))
//		}
//
//		cost, ok := costMap[productName]
//		if !ok {
//			cost = 0.0 // Если товар не найден в базе, устанавливаем себестоимость в 0
//		}
//		cost *= float64(quantity)
//
//		totalRevenue, err1 := parseFloat(rows[i][18]) // S в 19-й колонке (индексация с 0)
//		u, err2 := parseFloat(rows[i][20])            // U в 21-й колонке (индексация с 0)
//		aa, err3 := parseFloat(rows[i][26])           // AA в 27-й колонке (индексация с 0)
//
//		var ad float64
//		var err4 error
//		if rows[i][29] != "" {
//			ad, err4 = parseFloat(rows[i][29]) // AD в 30-й колонке (индексация с 0)
//		}
//
//		// Проверка на ошибки при преобразовании
//		if err1 != nil || err2 != nil || err3 != nil || (rows[i][29] != "" && err4 != nil) {
//			continue
//		}
//
//		sMinus3Percent := totalRevenue * 0.97
//		profit := sMinus3Percent + u + aa + ad - cost
//
//		newFile.SetCellValue(sheetName, fmt.Sprintf("AH%d", i+1), cost)
//		newFile.SetCellValue(sheetName, fmt.Sprintf("AI%d", i+1), profit)
//
//		totalProfit += profit
//	}
//
//	// Добавляем строку с общей чистой прибылью
//	summaryRow := len(rows) + 1
//	newFile.SetCellValue(sheetName, fmt.Sprintf("AH%d", summaryRow), "Общая Чистая Прибыль")
//	newFile.SetCellValue(sheetName, fmt.Sprintf("AI%d", summaryRow), totalProfit)
//
//	// Сохраняем обновленный файл
//	newFilePath := strings.Replace(filePath, ".xlsx", "_updated.xlsx", 1)
//	if err := newFile.SaveAs(newFilePath); err != nil {
//		return "", fmt.Errorf("не удалось сохранить файл: %v", err)
//	}
//
//	return newFilePath, nil
//}
//
//func main() {
//	bot, err := tgbotapi.NewBotAPI("7386513840:AAFzi2ixblH3DGeWTRoP0cURP6XP0A-ml3M")
//	if err != nil {
//		log.Panic(err)
//	}
//
//	bot.Debug = true
//
//	log.Printf("Authorized on account %s", bot.Self.UserName)
//
//	u := tgbotapi.NewUpdate(0)
//	u.Timeout = 60
//
//	updates := bot.GetUpdatesChan(u)
//
//	for update := range updates {
//		if update.Message == nil { // ignore any non-Message updates
//			continue
//		}
//
//		if update.Message.IsCommand() {
//			switch update.Message.Command() {
//			case "start":
//				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Привет! Отправьте мне файл себестоимости товаров (product_costs.csv).")
//				bot.Send(msg)
//				isCostFile = false
//				isReportFile = false
//			default:
//				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Я не знаю такой команды.")
//				bot.Send(msg)
//			}
//		}
//
//		if update.Message.Document != nil {
//			file := update.Message.Document
//			filePath := fmt.Sprintf("/tmp/%s", file.FileName)
//			fileURL, err := bot.GetFileDirectURL(file.FileID)
//			if err != nil {
//				msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при получении файла: %v", err))
//				bot.Send(msg)
//				continue
//			}
//
//			// Скачиваем файл
//			err = downloadFile(filePath, fileURL)
//			if err != nil {
//				msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при скачивании файла: %v", err))
//				bot.Send(msg)
//				continue
//			}
//
//			if strings.HasSuffix(file.FileName, ".csv") && !isCostFile {
//				// Загрузка себестоимости товаров
//				costMap, err = loadProductCosts(filePath)
//				if err != nil {
//					msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Не удалось загрузить себестоимость товаров: %v", err))
//					bot.Send(msg)
//					continue
//				}
//				isCostFile = true
//				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Файл себестоимости товаров загружен. Теперь отправьте Excel отчет для обработки.")
//				bot.Send(msg)
//			} else if strings.HasSuffix(file.FileName, ".xlsx") && isCostFile {
//				// Обработка файла
//				newFilePath, err := processExcel(filePath, costMap)
//				if err != nil {
//					msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при обработке файла: %v", err))
//					bot.Send(msg)
//					continue
//				}
//
//				// Отправка обработанного файла обратно
//				doc := tgbotapi.NewDocument(update.Message.Chat.ID, tgbotapi.FilePath(newFilePath))
//				doc.Caption = "Обработанный файл"
//				_, err = bot.Send(doc)
//				if err != nil {
//					msg := tgbotapi.NewMessage(update.Message.Chat.ID, fmt.Sprintf("Ошибка при отправке файла: %v", err))
//					bot.Send(msg)
//				}
//				isReportFile = true
//			} else {
//				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Пожалуйста, сначала отправьте файл себестоимости товаров (product_costs.csv), а затем Excel отчет (.xlsx).")
//				bot.Send(msg)
//			}
//		}
//	}
//}
//
//// Функция для скачивания файла по URL
////func downloadFile(filepath string, url string) error {
////	resp, err := http.Get(url)
////	if err != nil {
////		return err
////	}
////	defer resp.Body.Close()
////
////	out, err := os.Create(filepath)
////	if err != nil {
////		return err
////	}
////	defer out.Close()
////
////	_, err = io.Copy(out, resp.Body)
////	return err
////}
