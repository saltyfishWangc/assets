package com.example.msyd.assets;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

@SpringBootApplication
@EnableAutoConfiguration
public class AssetsApplication {

	public static ConfigurableApplicationContext ac;

	public static void main(String[] args) {
		ac = SpringApplication.run(AssetsApplication.class, args);
		Creator creator = (Creator) ac.getBean("creator");
		creator.create();
	}

}
