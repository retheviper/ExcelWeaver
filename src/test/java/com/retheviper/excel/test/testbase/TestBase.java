package com.retheviper.excel.test.testbase;

import com.retheviper.excel.test.header.Contract;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class TestBase {

    private static final DecimalFormat NUMBER_FORMAT = new DecimalFormat("0000");

    private static final DateFormat DATE_FORMAT = new SimpleDateFormat("yyyy/MM/dd");

    public static List<Contract> getTestData() {
        return IntStream.range(0, 10000).mapToObj(number -> {
            final String formatted = NUMBER_FORMAT.format(number);
            final Contract contract = new Contract();
            final String name = String.format("People%s", formatted);
            contract.setName(name);
            final String phone = String.format("(000) 0000-%s", formatted);
            contract.setCompanyPhone(phone);
            contract.setCellPhone(phone);
            contract.setHomePhone(phone);
            contract.setEmail(String.format("%s@example.com", name));
            contract.setBirth(createRandomDate());
            contract.setAddress("somewhere");
            contract.setCity("some city");
            contract.setProvince("somewhere");
            contract.setPost(90000 + number);
            return contract;
        }).collect(Collectors.toUnmodifiableList());
    }

    private static Date createRandomDate() {
        try {
            final long startMillis = DATE_FORMAT.parse("1970/01/01").getTime();
            final long endMillis = DATE_FORMAT.parse("2000/12/31").getTime();
            final long randomMillisSinceEpoch = ThreadLocalRandom
                    .current()
                    .nextLong(startMillis, endMillis);
            final Calendar calendar = Calendar.getInstance();
            calendar.setTime(new Date(randomMillisSinceEpoch));
            calendar.set(Calendar.HOUR, 0);
            calendar.set(Calendar.MINUTE, 0);
            calendar.set(Calendar.SECOND, 0);
            calendar.set(Calendar.HOUR_OF_DAY, 0);
            return calendar.getTime();
        } catch (ParseException e) {
            return new Date();
        }
    }
}
