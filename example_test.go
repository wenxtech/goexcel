package goexcel_test

import (
	"fmt"
	"sync"

	"github.com/wenxtech/goexcel"
)

type ExampleCost struct {
	Id    float64 `excel:"ID"`
	Name  string  `excel:"Name"`
	Email string  `excel:"Email"`
}

func ExampleReader() {
	r, err := goexcel.Reader("data.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer r.Close()
	var cost ExampleCost
	// r.AbleSheet([]string{"Sheet2", "Sheet5"})
	i := 0
	for {
		i++
		err = r.NextScan(&cost)
		if r.IsEnd() {
			break
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		fmt.Println(i, r.CurrSheetName(), r.CurrRowIndex(), cost.Id, cost.Name, cost.Email, err)
	}
}

func ExampleWriter() {
	w, err := goexcel.Writer()
	if err != nil {
		fmt.Println(err)
		return
	}
	defer w.Close()
	w.SetSheetMaxRow(10001)
	err = w.WriteHeader([]interface{}{"ID", "Name", "Email"})
	if err != nil {
		fmt.Println(err)
		return
	}
	var wg sync.WaitGroup

	for i := 0; i < 10; i++ {
		wg.Add(1)
		go func(startID int) {
			defer wg.Done()
			for j := 1; j <= 10000; j++ {
				id := startID*10000 + j
				data := []interface{}{id, fmt.Sprintf("User%d", id), fmt.Sprintf("user%d@example.com", id)}
				if err = w.WriteRow(data); err != nil {
					fmt.Println(err)
				}
			}

		}(i)
	}

	wg.Wait()

	if err = w.Save("data.xlsx"); err != nil {
		fmt.Println(err)
	}
}
