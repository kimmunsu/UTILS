package test;

public class TestExcel {
	public TestExcel(){};
	private String nickname;
	private short age;
	private String sex;
	public String getNickname() {
		return nickname;
	}
	public void setNickname(String nickname) {
		this.nickname = nickname;
	}
	public short getAge() {
		return age;
	}
	public void setAge(short age) {
		this.age = age;
	}
	public String getSex() {
		return sex;
	}
	public void setSex(String sex) {
		this.sex = sex;
	}
	@Override
	public String toString() {
		return "TestExcel [nickname=" + nickname + ", age=" + age + ", sex="
				+ sex + "]";
	}
}
